import os
import time
import tempfile
import requests
import schedule
import openpyxl
from datetime import datetime, date
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- AYARLAR ---
CREDENTIALS_FILE = r"C:\Users\VENTO\Desktop\credentials.json"
TOKEN_FILE        = r"C:\Users\VENTO\Desktop\token.json"
EXCEL_FILE        = r"C:\Users\VENTO\Desktop\post takip.xlsx"

PAGE_TOKEN     = os.environ.get("PAGE_TOKEN", "")
INSTAGRAM_ID   = os.environ.get("INSTAGRAM_ID", "17841477662272516")
PAGE_ID        = os.environ.get("PAGE_ID", "943024085570051")

DRIVE_ROOT     = "RA"
SCOPES         = ["https://www.googleapis.com/auth/drive.readonly"]


# --- GOOGLE DRIVE ---

def get_drive_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def find_folder(service, name, parent_id=None):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def find_file(service, name, parent_id):
    q = f"name='{name}' and '{parent_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def download_video(service, file_id, dest_path):
    req = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as f:
        dl = MediaIoBaseDownload(f, req, chunksize=10 * 1024 * 1024)
        done = False
        while not done:
            status, done = dl.next_chunk()
            print(f"  Drive: %{int(status.progress() * 100)} indirildi")


# --- INSTAGRAM ---

def upload_to_instagram(video_path, caption):
    print("  Instagram'a yukleniyor...")

    # 1. Resumable upload oturumu ac
    init = requests.post(
        f"https://graph.facebook.com/v24.0/{INSTAGRAM_ID}/media",
        data={
            "media_type": "REELS",
            "caption": caption,
            "upload_type": "resumable",
            "access_token": PAGE_TOKEN,
        }
    ).json()

    if "error" in init:
        print(f"  Instagram hata (oturum): {init['error']['message']}")
        return False

    upload_uri  = init.get("uri")
    creation_id = init.get("id")
    if not upload_uri or not creation_id:
        print(f"  Instagram: URI/ID alinamadi: {init}")
        return False

    # 2. Videoyu gonder
    file_size = os.path.getsize(video_path)
    with open(video_path, "rb") as f:
        video_bytes = f.read()

    upload_resp = requests.post(
        upload_uri,
        headers={
            "Authorization": f"OAuth {PAGE_TOKEN}",
            "offset": "0",
            "file_size": str(file_size),
        },
        data=video_bytes,
    )
    if upload_resp.status_code not in (200, 201):
        print(f"  Instagram yukleme hatasi ({upload_resp.status_code}): {upload_resp.text[:200]}")
        return False

    # 3. Islenmeyi bekle (max 5 dakika)
    print("  Instagram: video isleniyor...")
    for _ in range(30):
        time.sleep(10)
        status_resp = requests.get(
            upload_uri,
            headers={"Authorization": f"OAuth {PAGE_TOKEN}"},
            params={"fields": "upload_phase,status"},
        ).json()
        phase = status_resp.get("upload_phase", "")
        if phase == "finish":
            break
        if phase == "error":
            print(f"  Instagram: isleme hatasi: {status_resp}")
            return False

    # 4. Yayinla
    pub = requests.post(
        f"https://graph.facebook.com/v24.0/{INSTAGRAM_ID}/media_publish",
        data={"creation_id": creation_id, "access_token": PAGE_TOKEN},
    ).json()

    if "error" in pub:
        print(f"  Instagram hata (yayinla): {pub['error']['message']}")
        return False

    print(f"  Instagram: yuklendi! Post ID: {pub.get('id')}")
    return True


# --- FACEBOOK ---

def upload_to_facebook(video_path, caption):
    print("  Facebook'a yukleniyor...")

    with open(video_path, "rb") as f:
        resp = requests.post(
            f"https://graph.facebook.com/v24.0/{PAGE_ID}/videos",
            data={"description": caption, "access_token": PAGE_TOKEN},
            files={"source": f},
        ).json()

    if "error" in resp:
        print(f"  Facebook hata: {resp['error']['message']}")
        return False

    print(f"  Facebook: yuklendi! Video ID: {resp.get('id')}")
    return True


# --- EXCEL OKUMA ---

def parse_date(val):
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def parse_time(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    s = str(val).strip()
    # Excel zaman degeri (float 0.5 = 12:00)
    try:
        f = float(s)
        total_minutes = round(f * 24 * 60)
        h, m = divmod(total_minutes, 60)
        return f"{h:02d}:{m:02d}"
    except ValueError:
        return s[:5]  # "12:00:00" -> "12:00"


# --- ANA ISLEV ---

def process_posts():
    now = datetime.now()
    today = now.date()
    current_time = now.strftime("%H:%M")

    print(f"\n[{now.strftime('%Y-%m-%d %H:%M')}] Kontrol basliyor...")

    try:
        drive = get_drive_service()
        ra_id = find_folder(drive, DRIVE_ROOT)
        if not ra_id:
            print(f"Drive'da '{DRIVE_ROOT}' klasoru bulunamadi!")
            return
    except Exception as e:
        print(f"Drive baglanti hatasi: {e}")
        return

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    changed = False
    current_week = None

    for row in ws.iter_rows(min_row=2):
        # Hafta basligi satiri
        if row[0].value and "hafta" in str(row[0].value).lower():
            current_week = str(row[0].value).strip()
            continue

        video_name = row[1].value
        tarih      = parse_date(row[2].value)
        saat       = parse_time(row[3].value)
        aciklama   = str(row[4].value or "").strip()
        durum      = str(row[5].value or "").strip()

        if not video_name or durum.startswith("Yuklendi"):
            continue
        if tarih != today or saat != current_time:
            continue

        print(f"\nIslenecek: {video_name}  ({current_week})")

        # Drive'da hafta klasorunu bul
        week_id = find_folder(drive, current_week, ra_id)
        if not week_id:
            row[5].value = "Hata: hafta klasoru yok"
            changed = True
            continue

        reels_id = find_folder(drive, f"{current_week} Reelsler", week_id)
        if not reels_id:
            row[5].value = "Hata: Reelsler klasoru yok"
            changed = True
            continue

        filename = video_name if video_name.lower().endswith(".mp4") else f"{video_name}.mp4"
        file_id = find_file(drive, filename, reels_id)
        if not file_id:
            row[5].value = "Hata: video bulunamadi"
            changed = True
            continue

        # Videoyu gecici klasore indir
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        try:
            download_video(drive, file_id, temp_path)
        except Exception as e:
            print(f"  Indirme hatasi: {e}")
            row[5].value = f"Hata: {e}"
            changed = True
            continue

        ig_ok = upload_to_instagram(temp_path, aciklama)
        fb_ok = upload_to_facebook(temp_path, aciklama)

        try:
            os.remove(temp_path)
        except Exception:
            pass

        ts = now.strftime("%d.%m.%Y %H:%M")
        if ig_ok and fb_ok:
            row[5].value = f"Yuklendi - {ts}"
        elif ig_ok:
            row[5].value = f"Kismi: sadece Instagram - {ts}"
        elif fb_ok:
            row[5].value = f"Kismi: sadece Facebook - {ts}"
        else:
            row[5].value = f"Hata: her iki platform basarisiz - {ts}"

        changed = True

    if changed:
        wb.save(EXCEL_FILE)
        print("Excel guncellendi.")
    else:
        print("Bu saat icin bekleyen post bulunamadi.")


# --- CALISTIRMA ---

import sys

if "--once" in sys.argv:
    # GitHub Actions: tek seferlik calistir
    process_posts()
else:
    # Lokal: zamanlayici ile calistir
    schedule.every().day.at("12:00").do(process_posts)
    schedule.every().day.at("18:00").do(process_posts)
    print("Zamanlayici baslatildi: 12:00 ve 18:00")
    print("Durdurmak icin Ctrl+C\n")
    while True:
        schedule.run_pending()
        time.sleep(30)
