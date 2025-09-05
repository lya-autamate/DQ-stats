# run_report.py
import os
import sys
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
import papermill as pm
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError


# ========= НАСТРОЙКИ =========
# Путь к ноутбуку и директория вывода. Можно задать ENV NB_PATH/OUT_DIR.
NB_PATH = Path(os.environ.get("NB_PATH", "./notebooks/jira_report.ipynb")).resolve()
OUT_DIR = Path(os.environ.get("OUT_DIR", NB_PATH.parent)).resolve()

# Таймзона для имен файлов (советуем Asia/Tashkent, чтобы совпадало с ноутбуком)
TZ = os.getenv("TZ", "Asia/Tashkent")
now = datetime.now(ZoneInfo(TZ))
today_str = now.strftime("%Y-%m-%d")     # для xlsx
month_str = now.strftime("%B %Y")        # для pdf, напр. "September 2025"

# Slack
load_dotenv()
SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_CHANNEL_NAME = os.getenv("SLACK_CHANNEL_NAME", "dq_chat")

if not SLACK_BOT_TOKEN:
    print("ERROR: SLACK_BOT_TOKEN не задан.", file=sys.stderr)
    sys.exit(1)

# ========= ШАГ 1. Запускаем ноутбук =========
print(f"Executing notebook: {NB_PATH}")
OUT_DIR.mkdir(parents=True, exist_ok=True)
executed_nb = OUT_DIR / "executed.ipynb"

pm.execute_notebook(
    input_path=str(NB_PATH),
    output_path=str(executed_nb),
    cwd=str(NB_PATH.parent),   # чтобы относительные пути внутри ноутбука совпадали с ручным запуском
    parameters={},             # при необходимости передайте параметры ноутбуку
)

# ========= ШАГ 2. Ищем файлы с точными именами =========
xlsx_path = OUT_DIR / f"выгрузка из JIRA ({today_str}).xlsx"
pdf_path  = OUT_DIR / f"jira_report ({month_str}).pdf"

def ensure_file_exists(p: Path):
    if not p.exists():
        # небольшая помощь в отладке
        print(f"ERROR: файл не найден: {p}", file=sys.stderr)
        print("Содержимое OUT_DIR:", OUT_DIR, file=sys.stderr)
        for fp in sorted(OUT_DIR.glob("*")):
            print(" -", fp.name, file=sys.stderr)
        sys.exit(2)

ensure_file_exists(xlsx_path)
ensure_file_exists(pdf_path)

print(f"Found files:\n - {xlsx_path.name}\n - {pdf_path.name}")

# ========= ШАГ 3. Готовим Slack-клиент =========
client = WebClient(token=SLACK_BOT_TOKEN)

def resolve_channel_id(name: str) -> str:
    cursor = None
    for _ in range(20):
        resp = client.conversations_list(
            exclude_archived=True,
            limit=1000,
            types="public_channel,private_channel",
            cursor=cursor,
        )
        for ch in resp["channels"]:
            if ch.get("name") == name or ch.get("name_normalized") == name:
                return ch["id"]
        cursor = resp.get("response_metadata", {}).get("next_cursor") or None
        if not cursor:
            break
    raise RuntimeError(f"Канал '{name}' не найден. Проверь имя и что бот добавлен в канал.")

channel_id = resolve_channel_id(SLACK_CHANNEL_NAME)

# ========= ШАГ 4. Отправляем файлы =========
comment = f"Ежемесячный отчёт: {month_str}"

def upload_file(path: Path, with_comment: bool = False):
    try:
        client.files_upload_v2(
            channel=channel_id,
            filename=path.name,
            file=path.open("rb"),
            title=path.name,
            initial_comment=(comment if with_comment else None),
        )
        print(f"Uploaded: {path.name}")
    except SlackApiError as e:
        print(f"Slack error for {path.name}: {e.response.get('error')}", file=sys.stderr)
        raise

# Сначала Excel (без комментария), затем PDF (с комментарием)
upload_file(xlsx_path, with_comment=False)
upload_file(pdf_path,  with_comment=True)

print("Done.")