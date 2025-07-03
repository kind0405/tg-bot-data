import os
import json
import io
import docx
import PyPDF2
import openpyxl
import requests

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- Настройки ---
SERVICE_ACCOUNT_FILE = 'credentials.json'
FOLDER_ID = '1SDBfV-2Zk7lriKUsgRSS6wWnyC2O7ZX0'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly',
          'https://www.googleapis.com/auth/documents.readonly']
CLOUDFLARE_API_TOKEN = os.getenv('CLOUDFLARE_API_TOKEN')  # должен быть установлен как переменная окружения

CLOUDFLARE_ACCOUNT_ID = '215a19df1120e2c3a787071fa5a05dd9'
EMBEDDING_MODEL = '@cf/baai/bge-base-ru-v1.5'
EMBEDDING_ENDPOINT = f'https://api.cloudflare.com/client/v4/accounts/{CLOUDFLARE_ACCOUNT_ID}/ai/run/{EMBEDDING_MODEL}'

SUPPORTED_MIME_TYPES = [
    'application/vnd.google-apps.document',
    'application/vnd.google-apps.spreadsheet',
    'application/pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
]

# --- Авторизация ---
def get_services():
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build('drive', 'v3', credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)
    return drive_service, docs_service

# --- Получение списка файлов в папке ---
def list_files(service, folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    files = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType)',
            pageToken=page_token
        ).execute()
        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    return files

# --- Загрузка бинарных файлов ---
def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# --- Google Docs ---
def extract_text_from_google_doc(docs_service, file_id):
    doc = docs_service.documents().get(documentId=file_id).execute()
    text = []
    for element in doc.get('body', {}).get('content', []):
        paragraph = element.get('paragraph')
        if paragraph:
            for elem in paragraph.get('elements', []):
                text_run = elem.get('textRun')
                if text_run:
                    text.append(text_run.get('content', ''))
    return ''.join(text)

# --- DOCX ---
def extract_text_from_docx(file_io):
    doc = docx.Document(file_io)
    return '\n'.join([p.text for p in doc.paragraphs])

# --- PDF ---
def extract_text_from_pdf(file_io):
    reader = PyPDF2.PdfReader(file_io)
    return '\n'.join([page.extract_text() or '' for page in reader.pages])

# --- XLSX ---
def extract_text_from_xlsx(file_io):
    wb = openpyxl.load_workbook(file_io)
    text = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            text.append(' '.join([str(cell) if cell else '' for cell in row]))
    return '\n'.join(text)

# --- Получение embedding через Cloudflare AI ---
def get_embedding(text, api_token):
    headers = {
        "Authorization": f"Bearer {api_token}",
        "Content-Type": "application/json"
    }
    body = {
        "text": text
    }
    try:
        response = requests.post(EMBEDDING_ENDPOINT, headers=headers, json=body)
        if response.status_code == 200:
            return response.json()['result']['data']
        else:
            print("⚠️ Ошибка от Cloudflare AI:", response.status_code, response.text)
            return []
    except Exception as e:
        print("⚠️ Исключение при запросе embedding:", e)
        return []

# --- Главная функция ---
def main():
    drive_service, docs_service = get_services()
    files = list_files(drive_service, FOLDER_ID)
    data = []

    for file in files:
        file_id = file['id']
        name = file['name']
        mime_type = file['mimeType']
        print(f"📄 Обрабатываем: {name} ({file_id})")

        content = "[Не поддерживается]"

        try:
            if mime_type == 'application/vnd.google-apps.document':
                content = extract_text_from_google_doc(docs_service, file_id)

            elif mime_type == 'application/vnd.google-apps.spreadsheet':
                request = drive_service.files().export_media(fileId=file_id,
                    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                content = extract_text_from_xlsx(fh)

            elif mime_type == 'application/pdf':
                fh = download_file(drive_service, file_id)
                content = extract_text_from_pdf(fh)

            elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                fh = download_file(drive_service, file_id)
                content = extract_text_from_docx(fh)

            elif mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                fh = download_file(drive_service, file_id)
                content = extract_text_from_xlsx(fh)

            else:
                print(f"⛔️ Формат не поддерживается: {mime_type}")

        except Exception as e:
            print(f"❌ Ошибка при обработке {name}: {e}")

        # Генерируем embedding только если есть содержимое
        if content != "[Не поддерживается]" and CLOUDFLARE_API_TOKEN:
            embedding = get_embedding(content[:1000], CLOUDFLARE_API_TOKEN)
        else:
            embedding = []

        data.append({
            'id': file_id,
            'name': name,
            'mimeType': mime_type,
            'content': content.strip(),
            'embedding': embedding
        })

    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("✅ Файл data.json успешно создан!")

if __name__ == '__main__':
    main()