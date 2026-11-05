#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ОТТС Downloader Pro v6
Исправлена ошибка IllegalCharacterError в openpyxl: добавлена очистка строк
от недопустимых управляющих символов перед записью в Excel.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import requests
import json
import time
import os
import re
import glob
import shutil
import threading
import logging
from datetime import datetime
import urllib3
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= КОНСТАНТЫ =================
API_URL = "https://www.gost.ru/standards/rest-api/libraries/search"
FILE_SERVICE_URL = "https://rst.gov.ru:8443/file-service/file/load/"
CONFIG_FILE = "otts_config.json"

CATALOG_DIR = "catalog"
CATALOG_JS_DIR = os.path.join(CATALOG_DIR, "js")
CATALOG_PROGRESS_FILE = os.path.join(CATALOG_DIR, "catalog_progress.json")

PAGE_SIZE = 20
REQUEST_DELAY = 3.0
MAX_RETRIES = 5
CONNECT_TIMEOUT = 30
READ_TIMEOUT = 240
EMPTY_PLACEHOLDER = "––"

# Регулярка для удаления недопустимых в XML/Excel символов
# (ASCII control characters: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F)
ILLEGAL_CHARACTERS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

DEFAULT_HEADERS = {
    "Accept": "application/json",
    "gtn.access.appid": "ВСТАВЬТЕ_СЮДА_СВОЙ_APPID",
    "App-Id": "ВСТАВЬТЕ_СЮДА_СВОЙ_APP_ID",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.gost.ru/portal/gost/home/activity/compliance/evaluationcompliance/AcknowledgementCorrespondence/vehicletypeapproval018",
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:140.0) Gecko/20100101 Firefox/140.0",
    "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3"
}

XLSX_COLUMNS = [
    "№", "ID",
    "Наименование органа по сертификации",
    "Номер одобрения типа по реестру",
    "Дата регистрации одобрения типа",
    "Марка",
    "Коммерческое наименование",
    "Тип",
    "Модификации",
    "Шасси",
    "Модификации (таблица)",
    "Категория",
    "Код VIN",
    "Экологический класс",
    "Двигатель",
    "Изготовитель и его адрес",
    "Сборочный(е) завод(ы) и его(их) адрес(а)",
    "Файл с текстом одобрения типа",
    "Номер одобрения типа (ссылка на сайт)"
]
COL_WIDTHS = [6, 8, 40, 30, 14, 20, 30, 25, 25, 20, 25, 10, 30, 12, 30, 50, 50, 30, 40]

COL_NUM = 1
COL_ID = 2
COL_CERT_NUM = 4
COL_DATE = 5
COL_PDF = 18
COL_SITE = 19


def sanitize_for_excel(value):
    """
    Очищает строку от недопустимых в Excel/XML символов.
    - Заменяет неразрывные пробелы (\xa0) на обычные.
    - Удаляет управляющие ASCII-символы.
    - Нормализует переносы строк (\r\n -> \n).
    """
    if not isinstance(value, str):
        return value
    # Неразрывный пробел → обычный
    value = value.replace('\xa0', ' ')
    # CRLF → LF
    value = value.replace('\r\n', '\n').replace('\r', '\n')
    # Удаляем недопустимые управляющие символы
    value = ILLEGAL_CHARACTERS_RE.sub('', value)
    return value


def create_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": DEFAULT_HEADERS["User-Agent"],
        "Accept-Language": DEFAULT_HEADERS["Accept-Language"],
        "Accept": "application/json"
    })
    adapter = requests.adapters.HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=0)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


# ============================================================
#  ФАЗА 1: Каталог
# ============================================================
class CatalogDownloader:
    def __init__(self, session_cookie, logger, progress_callback=None):
        self.session_cookie = session_cookie
        self.logger = logger
        self.progress_callback = progress_callback
        self._session = create_session()
        self.stop_flag = False
        self.stats = {"total_pages": 0, "current_page": 0, "skipped": 0, "downloaded": 0, "errors": 0}

    def get_headers(self):
        h = DEFAULT_HEADERS.copy()
        h["Cookie"] = f"session-cookie={self.session_cookie}; sp_test=1"
        return h

    def load_local_page(self, page):
        path = os.path.join(CATALOG_JS_DIR, f"page_{page}.json")
        if not os.path.exists(path):
            return None
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if "data" in data:
                return data
        except Exception:
            pass
        return None

    def migrate_old_json(self):
        moved = 0
        for folder in glob.glob("Реестр ОТТС_*"):
            js_dir = os.path.join(folder, "js")
            if not os.path.isdir(js_dir):
                continue
            for pf in glob.glob(os.path.join(js_dir, "page_*.json")):
                dst = os.path.join(CATALOG_JS_DIR, os.path.basename(pf))
                if not os.path.exists(dst):
                    try:
                        shutil.copy(pf, dst)
                        moved += 1
                    except Exception as e:
                        self.logger.warning(f"Не удалось скопировать {pf}: {e}")
        if moved:
            self.logger.info(f"✓ Миграция: скопировано {moved} JSON из старых папок")

    def fetch_and_save(self, page):
        params = {"query": "", "page": page, "size": PAGE_SIZE, "tmsp": int(time.time() * 1000)}
        last_error = None
        for attempt in range(MAX_RETRIES):
            try:
                self.logger.info(f"  Запрос страницы {page} (попытка {attempt + 1}/{MAX_RETRIES})...")
                response = self._session.get(API_URL, params=params, headers=self.get_headers(),
                                             verify=False, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT))
                if response.status_code == 401:
                    raise Exception("Ошибка 401: Куки устарели! Обновите session-cookie.")
                if response.status_code == 429:
                    wait = 60 * (attempt + 1)
                    self.logger.warning(f"  ⚠ 429. Ждем {wait}с...")
                    time.sleep(wait)
                    continue
                if response.status_code != 200:
                    raise Exception(f"HTTP {response.status_code}")
                data = response.json()
                with open(os.path.join(CATALOG_JS_DIR, f"page_{page}.json"), 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                return data
            except requests.exceptions.RequestException as e:
                last_error = e
                wait = 30 * (attempt + 1)
                self.logger.warning(f"  ⚠ {type(e).__name__}. Ждем {wait}с...")
                time.sleep(wait)
        self.stats["errors"] += 1
        raise Exception(f"Страница {page} не получена: {last_error}")

    def run(self):
        try:
            os.makedirs(CATALOG_JS_DIR, exist_ok=True)
            self.migrate_old_json()
            self.logger.info("Определяем объем базы...")
            page1 = self.load_local_page(1)
            if page1 is None:
                page1 = self.fetch_and_save(1)
                self.stats["downloaded"] += 1
            total_count = page1.get("count", 0)
            total_pages = (total_count + PAGE_SIZE - 1) // PAGE_SIZE
            self.stats["total_pages"] = total_pages
            self.logger.info(f"Всего документов: {total_count}, страниц: {total_pages}")
            for page in range(1, total_pages + 1):
                if self.stop_flag:
                    break
                self.stats["current_page"] = page
                if self.load_local_page(page) is not None:
                    self.stats["skipped"] += 1
                else:
                    self.fetch_and_save(page)
                    self.stats["downloaded"] += 1
                    time.sleep(REQUEST_DELAY)
                if self.progress_callback:
                    self.progress_callback(self.stats, page, total_pages)
            if not self.stop_flag:
                with open(CATALOG_PROGRESS_FILE, 'w', encoding='utf-8') as f:
                    json.dump({"completed": True, "total_pages": total_pages,
                               "timestamp": datetime.now().isoformat()}, f, ensure_ascii=False, indent=2)
                self.logger.info(f"\n✅ КАТАЛОГ ГОТОВ! Страниц: {total_pages}")
            else:
                self.logger.info("\n⏸ Каталог прерван.")
            return True
        except Exception as e:
            self.logger.error(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False


# ============================================================
#  ФАЗА 2: Обработка года (v6)
# ============================================================
class YearProcessor:
    STATUS_DONE = "done"
    STATUS_NO_PDF = "no_pdf"
    STATUS_ERROR = "error"
    STATUS_LEGACY = "legacy"

    def __init__(self, year, session_cookie, logger, progress_callback=None):
        self.year = year
        self.session_cookie = session_cookie
        self.logger = logger
        self.progress_callback = progress_callback
        self._session = create_session()
        self.stop_flag = False

        self.base_dir = f"Реестр ОТТС_{year}" if year != "all" else "Реестр ОТТС_все_годы"
        self.pdf_dir = os.path.join(self.base_dir, "pdf")
        self.xlsx_path = os.path.join(self.base_dir, f"otts_{year}.xlsx")
        self.processed_file = os.path.join(self.base_dir, "processed_ids.json")

        self.stats = {
            "total_pages": 0, "current_page": 0, "found": 0,
            "downloaded": 0, "existing_reused": 0, "no_pdf": 0, "errors": 0
        }
        self.existing_files = set()
        self.processed = {}
        self.xlsx_row_by_id = {}
        self.used_pdf_names = set()

    def get_headers(self):
        h = DEFAULT_HEADERS.copy()
        h["Cookie"] = f"session-cookie={self.session_cookie}; sp_test=1"
        return h

    def load_processed(self):
        if not os.path.exists(self.processed_file):
            self.processed = {}
            return
        try:
            with open(self.processed_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                self.logger.warning(
                    f"Старый формат processed_ids.json ({len(data)} ID) — "
                    f"помечены как 'legacy' и будут проверены повторно."
                )
                self.processed = {int(x): self.STATUS_LEGACY for x in data if x is not None}
            elif isinstance(data, dict):
                self.processed = {int(k): v for k, v in data.items()}
            else:
                self.processed = {}
        except Exception as e:
            self.logger.warning(f"Не удалось загрузить processed_ids.json: {e}")
            self.processed = {}

    def save_processed(self):
        with open(self.processed_file, 'w', encoding='utf-8') as f:
            json.dump({str(k): v for k, v in self.processed.items()}, f, ensure_ascii=False, indent=2)

    def should_skip(self, doc_id):
        return self.processed.get(doc_id) == self.STATUS_DONE

    def extract_year(self, date_str):
        if not date_str:
            return None
        match = re.search(r'(19|20)\d{2}', str(date_str))
        if match:
            year = int(match.group(0))
            if 1900 <= year <= 2100:
                return year
        return None

    def safe_filename(self, text):
        if not text:
            return "unknown"
        safe = re.sub(r'[\\/*?:"<>|]', "_", str(text))
        safe = safe.replace(" ", "_")
        return safe[:100]

    def extract_field(self, fields, key, default=""):
        value = fields.get(key)
        return default if value is None else str(value).strip()

    def fill_empty(self, value):
        if value is None:
            return EMPTY_PLACEHOLDER
        s = str(value).strip()
        if s == "" or s == "-":
            return EMPTY_PLACEHOLDER
        return s

    def parse_date(self, date_str):
        if not date_str:
            return None
        date_part = str(date_str).strip().split()[0]
        try:
            return datetime.strptime(date_part, "%d.%m.%Y")
        except Exception:
            return None

    def find_existing_pdf(self, safe_cert):
        exact = f"{safe_cert}.pdf"
        if exact in self.existing_files:
            return exact
        prefix = f"{safe_cert}_"
        for f in self.existing_files:
            if f.startswith(prefix) and f.lower().endswith(".pdf"):
                return f
        return None

    def download_pdf(self, url, filename):
        filepath = os.path.join(self.pdf_dir, filename)
        existing = self.find_existing_pdf(filename[:-4])
        if existing:
            return True, True, existing

        for attempt in range(MAX_RETRIES):
            try:
                response = self._session.get(url, verify=False,
                                             timeout=(CONNECT_TIMEOUT, READ_TIMEOUT), stream=True)
                if response.status_code != 200:
                    raise Exception(f"HTTP {response.status_code}")
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                self.existing_files.add(filename)
                return True, False, filename
            except Exception as e:
                if attempt < MAX_RETRIES - 1:
                    time.sleep(2 * (attempt + 1))
                else:
                    self.logger.error(f"  ✗ Ошибка скачивания {filename}: {e}")
                    return False, False, ""
        return False, False, ""

    def fetch_pdf_link_from_card(self, card_url):
        if not card_url:
            return None
        try:
            response = self._session.get(card_url, headers=self.get_headers(),
                                         verify=False, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT))
            if response.status_code != 200:
                return None
            match = re.search(r'file/load/(\d{13})', response.text)
            if match:
                return match.group(1)
        except Exception as e:
            self.logger.warning(f"  ⚠ Не удалось получить карточку: {e}")
        return None

    def create_or_open_xlsx(self):
        if os.path.exists(self.xlsx_path):
            wb = load_workbook(self.xlsx_path)
            ws = wb.active
            self.logger.info(f"Открыт существующий XLSX, строк: {ws.max_row}")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = f"ОТТС {self.year}"
            for col_num, header in enumerate(XLSX_COLUMNS, 1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(wrap_text=True, vertical='top')
            for i, w in enumerate(COL_WIDTHS):
                col_letter = chr(65 + i) if i < 26 else 'A' + chr(65 + i - 26)
                ws.column_dimensions[col_letter].width = w
            ws.column_dimensions['B'].hidden = True
            ws.freeze_panes = 'A2'

        self.xlsx_row_by_id = {}
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=COL_ID).value
            if val is not None:
                try:
                    self.xlsx_row_by_id[int(val)] = row
                except (ValueError, TypeError):
                    pass

        return wb, ws

    def write_row(self, ws, doc_id, row_data):
        if doc_id in self.xlsx_row_by_id:
            row = self.xlsx_row_by_id[doc_id]
            action = "updated"
        else:
            row = ws.max_row + 1
            self.xlsx_row_by_id[doc_id] = row
            action = "added"

        for col_num, value in enumerate(row_data, 1):
            # === ИСПРАВЛЕНИЕ v6: санитизация значения перед записью ===
            value = sanitize_for_excel(value)
            cell = ws.cell(row=row, column=col_num, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical='top')

            if col_num == COL_DATE and isinstance(value, datetime):
                cell.number_format = 'DD.MM.YYYY'

        pdf_cell = ws.cell(row=row, column=COL_PDF)
        pdf_value = pdf_cell.value
        if pdf_value and isinstance(pdf_value, str) and pdf_value.strip() and pdf_value != EMPTY_PLACEHOLDER:
            pdf_cell.hyperlink = f"pdf/{pdf_value}"
            pdf_cell.font = Font(color="0000FF", underline="single")

        return row, action

    def set_site_hyperlink(self, ws, row, card_url):
        if card_url:
            cell = ws.cell(row=row, column=COL_SITE)
            cell.hyperlink = card_url
            cell.font = Font(color="0000FF", underline="single")

    def process_document(self, doc, ws):
        doc_id = doc.get("id")
        fields = doc.get("fields", {})
        attachments = doc.get("attachmentDtoList", [])

        pub_date_raw = self.extract_field(fields, "dateofpublication")
        doc_year = self.extract_year(pub_date_raw)

        if self.year != "all" and doc_year != int(self.year):
            return "skipped_year"

        cert_number_raw = self.extract_field(fields, "numberofcertificate") or f"[без номера, id={doc_id}]"
        self.logger.info(f"\n  📄 [{doc_id}] {cert_number_raw} (год: {doc_year})")

        cert_org = self.extract_field(fields, "nameofcertificationorg")
        product = self.extract_field(fields, "product")
        tradename = self.extract_field(fields, "tradename")
        vtype = self.extract_field(fields, "type")
        modifications = self.extract_field(fields, "modifications")
        chassis = self.extract_field(fields, "chassis")
        mod_table = self.extract_field(fields, "modificationstable11")
        category = self.extract_field(fields, "category")
        codevin = self.extract_field(fields, "codevin")
        ecoclass = self.extract_field(fields, "ecoclass")
        engine = self.extract_field(fields, "enginetrademark")

        manufacturer = self.extract_field(fields, "nameofmanufacturer")
        m_address = self.extract_field(fields, "addressofmanufacturer")
        manufacturer_full = ""
        if manufacturer or m_address:
            parts = [p for p in [manufacturer, m_address] if p and p.strip() and p.strip() != "-"]
            manufacturer_full = "\n".join(parts)

        plants = []
        for i in range(1, 9):
            plant = self.extract_field(fields, f"assemblyplant_{i}")
            if plant and plant.strip() and plant.strip() != "-":
                plants.append(plant)
        plants_str = "\n\n".join(plants)

        card_url = doc.get("url", "")

        pdf_filename = ""
        download_url = None
        uuid = None

        pdf_attachments = [a for a in attachments if a.get("extension", "").lower() == "pdf"]

        if pdf_attachments:
            attachment = pdf_attachments[0]
            uuid = attachment.get("uniqId") or attachment.get("uuid")
            download_url = attachment.get("downloadLink") or f"{FILE_SERVICE_URL}{uuid}"
            self.logger.info(f"     ✓ Вложение найдено в каталоге (uuid={uuid})")
        else:
            self.logger.info(f"     ℹ В каталоге вложений нет — иду в карточку...")
            uuid = self.fetch_pdf_link_from_card(card_url)
            if uuid:
                download_url = f"{FILE_SERVICE_URL}{uuid}"
                self.logger.info(f"     ✓ Ссылка найдена через карточку (uuid={uuid})")
            else:
                self.logger.info(f"     ⚠ PDF не найден нигде")
            time.sleep(0.5)

        pdf_status = self.STATUS_NO_PDF
        actual_pdf_name = ""
        if download_url and uuid:
            safe_cert = self.safe_filename(cert_number_raw) if cert_number_raw and not cert_number_raw.startswith("[") else f"doc_{doc_id}"
            pdf_filename = f"{safe_cert}.pdf"

            if pdf_filename in self.used_pdf_names:
                counter = 2
                while f"{safe_cert}_{counter}.pdf" in self.used_pdf_names:
                    counter += 1
                pdf_filename = f"{safe_cert}_{counter}.pdf"
                self.logger.warning(f"     ⚠ Конфликт имен: использовано {pdf_filename}")

            success, reused, actual_pdf_name = self.download_pdf(download_url, pdf_filename)
            if success:
                self.used_pdf_names.add(actual_pdf_name)
                if reused:
                    self.stats["existing_reused"] += 1
                    self.logger.info(f"     ♻ PDF уже был в папке: {actual_pdf_name}")
                else:
                    self.stats["downloaded"] += 1
                    self.logger.info(f"     ⬇ Скачан новый PDF: {actual_pdf_name}")
                pdf_status = self.STATUS_DONE
            else:
                self.stats["errors"] += 1
                pdf_status = self.STATUS_ERROR
                actual_pdf_name = ""
        else:
            self.stats["no_pdf"] += 1

        parsed_date = self.parse_date(pub_date_raw)

        row_data = [
            None, doc_id,
            self.fill_empty(cert_org),
            self.fill_empty(cert_number_raw),
            parsed_date if parsed_date else EMPTY_PLACEHOLDER,
            self.fill_empty(product),
            self.fill_empty(tradename),
            self.fill_empty(vtype),
            self.fill_empty(modifications),
            self.fill_empty(chassis),
            self.fill_empty(mod_table),
            self.fill_empty(category),
            self.fill_empty(codevin),
            self.fill_empty(ecoclass),
            self.fill_empty(engine),
            self.fill_empty(manufacturer_full),
            self.fill_empty(plants_str),
            actual_pdf_name if actual_pdf_name else EMPTY_PLACEHOLDER,
            self.fill_empty(cert_number_raw)
        ]

        row, action = self.write_row(ws, doc_id, row_data)
        self.set_site_hyperlink(ws, row, card_url)
        self.logger.info(f"     📝 Строка XLSX {row} ({action})")

        return pdf_status

    def renumber_xlsx(self, ws):
        count = 0
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=COL_ID).value is not None:
                count += 1
                ws.cell(row=row, column=COL_NUM, value=count)
        self.logger.info(f"  ✓ Нумерация проставлена: 1..{count}")
        return count

    def run(self):
        try:
            page_files = sorted(
                glob.glob(os.path.join(CATALOG_JS_DIR, "page_*.json")),
                key=lambda p: int(re.search(r'page_(\d+)\.json', p).group(1))
            )
            if not page_files:
                self.logger.error("❌ Каталог пуст! Сначала скачайте каталог.")
                return False

            os.makedirs(self.pdf_dir, exist_ok=True)
            self.existing_files = set(f for f in os.listdir(self.pdf_dir) if f.lower().endswith(".pdf"))
            self.used_pdf_names = set(self.existing_files)
            self.logger.info(f"В папке pdf/ уже есть {len(self.existing_files)} файлов")

            self.load_processed()
            done_count = sum(1 for s in self.processed.values() if s == self.STATUS_DONE)
            legacy_count = sum(1 for s in self.processed.values() if s == self.STATUS_LEGACY)
            recheck_count = sum(1 for s in self.processed.values() if s in (self.STATUS_NO_PDF, self.STATUS_ERROR))
            self.logger.info(
                f"Прогресс: done={done_count}, legacy={legacy_count}, "
                f"no_pdf/error={recheck_count} (будут перепроверены)"
            )

            wb, ws = self.create_or_open_xlsx()
            self.stats["total_pages"] = len(page_files)
            self.logger.info(f"Обрабатываем {len(page_files)} страниц каталога для года: {self.year}\n")

            for idx, pf in enumerate(page_files, 1):
                if self.stop_flag:
                    break
                self.stats["current_page"] = idx

                try:
                    with open(pf, 'r', encoding='utf-8') as f:
                        docs = json.load(f).get("data", [])
                except Exception as e:
                    self.logger.warning(f"⚠ Пропускаем {pf}: {e}")
                    self.stats["errors"] += 1
                    continue

                page_found = 0
                page_done = 0
                page_no_pdf = 0
                for doc in docs:
                    if self.stop_flag:
                        break
                    doc_id = doc.get("id")
                    if doc_id is None:
                        continue
                    if self.should_skip(doc_id):
                        continue

                    try:
                        status = self.process_document(doc, ws)
                    except Exception as e:
                        # Локальная защита: если один документ падает — не валим всю загрузку
                        self.logger.error(f"     ✗ Ошибка обработки документа {doc_id}: {e}")
                        import traceback
                        self.logger.error(traceback.format_exc())
                        status = self.STATUS_ERROR
                        self.stats["errors"] += 1

                    if status == "skipped_year":
                        continue

                    page_found += 1
                    self.stats["found"] += 1
                    self.processed[doc_id] = status

                    if status == self.STATUS_DONE:
                        page_done += 1
                    elif status == self.STATUS_NO_PDF:
                        page_no_pdf += 1

                wb.save(self.xlsx_path)
                self.save_processed()

                if page_found > 0:
                    self.logger.info(
                        f"\n--- Стр. {idx}/{len(page_files)} --- "
                        f"За {self.year} год: +{page_found} док. | "
                        f"Успешно: {page_done} | Без PDF: {page_no_pdf}"
                    )
                if self.progress_callback:
                    self.progress_callback(self.stats, idx, len(page_files))

            self.logger.info("\nПересчет нумерации строк...")
            self.renumber_xlsx(ws)
            wb.save(self.xlsx_path)

            self.logger.info(f"\n{'=' * 60}")
            self.logger.info(f"✅ ГОТОВО за {self.year}!")
            self.logger.info(f"Всего документов за год: {self.stats['found']}")
            self.logger.info(f"  • Скачано новых PDF: {self.stats['downloaded']}")
            self.logger.info(f"  • Переиспользовано из папки: {self.stats['existing_reused']}")
            self.logger.info(f"  • Без PDF (нет нигде): {self.stats['no_pdf']}")
            self.logger.info(f"  • Ошибок скачивания: {self.stats['errors']}")
            self.logger.info(f"XLSX: {self.xlsx_path}")
            self.logger.info(f"{'=' * 60}")
            return True

        except Exception as e:
            self.logger.error(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False


# ============================================================
#  GUI
# ============================================================
class OTTSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ОТТС Downloader Pro v6")
        self.root.geometry("1000x780")
        self.root.minsize(850, 620)
        self.config = self.load_config()
        self.worker = None
        self.thread = None
        self.setup_logging()
        self.build_ui()
        self.refresh_catalog_status()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {"session_cookie": ""}

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def setup_logging(self):
        self.logger = logging.getLogger("OTTS")
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)
                self.text_widget.after(0, self.append, msg)

            def append(self, msg):
                self.text_widget.configure(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')

        fh = logging.FileHandler('otts_downloader.log', encoding='utf-8')
        fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(fh)
        self._text_handler_class = TextHandler

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        settings_frame = ttk.LabelFrame(main_frame, text="⚙ Настройки", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(settings_frame, text="Session Cookie:").grid(row=0, column=0, sticky="w", pady=5)
        self.cookie_var = tk.StringVar(value=self.config.get("session_cookie", ""))
        ttk.Entry(settings_frame, textvariable=self.cookie_var, width=80).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(settings_frame, text="💾 Сохранить", command=self.save_cookie).grid(row=0, column=2, padx=5, pady=5)
        settings_frame.columnconfigure(1, weight=1)

        catalog_frame = ttk.LabelFrame(main_frame, text="📚 Фаза 1 — Каталог (делается ОДИН раз)", padding=10)
        catalog_frame.pack(fill=tk.X, pady=(0, 10))
        self.catalog_status_var = tk.StringVar(value="")
        ttk.Label(catalog_frame, textvariable=self.catalog_status_var, font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        self.btn_catalog = ttk.Button(catalog_frame, text="📥 Скачать каталог (JSON)", command=self.start_catalog)
        self.btn_catalog.pack(side=tk.RIGHT, padx=5)

        year_frame = ttk.LabelFrame(main_frame, text="📅 Фаза 2 — Обработка года", padding=10)
        year_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(year_frame, text="Год:").pack(side=tk.LEFT, padx=5)
        years = ["Все годы"] + [str(y) for y in range(2015, 2027)]
        self.year_var = tk.StringVar(value="2015")
        ttk.Combobox(year_frame, textvariable=self.year_var, values=years, state="readonly", width=12).pack(side=tk.LEFT, padx=5)
        self.btn_year = ttk.Button(year_frame, text="▶ Обработать выбранный год", command=self.start_year)
        self.btn_year.pack(side=tk.LEFT, padx=10)
        self.btn_reset_year = ttk.Button(year_frame, text="🗑 Сбросить год", command=self.reset_year)
        self.btn_reset_year.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(year_frame, text="⏸ Остановить", command=self.stop_worker, state="disabled")
        self.btn_stop.pack(side=tk.RIGHT, padx=5)

        stats_frame = ttk.LabelFrame(main_frame, text="📊 Статистика", padding=10)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(stats_frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=5)
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X)
        self.caption_labels = {}
        self.stats_labels = {}
        stat_items = [
            ("total", "Всего страниц:"), ("found", "Найдено за год:"),
            ("downloaded", "Скачано PDF:"), ("errors", "Ошибок:"), ("page", "Страница:")
        ]
        for i, (key, caption) in enumerate(stat_items):
            cap = ttk.Label(stats_grid, text=caption, font=("Arial", 10))
            cap.grid(row=0, column=i * 2, sticky="w", padx=5, pady=2)
            self.caption_labels[key] = cap
            val = ttk.Label(stats_grid, text="0", font=("Arial", 10, "bold"))
            val.grid(row=0, column=i * 2 + 1, sticky="w", padx=2, pady=2)
            self.stats_labels[key] = val

        log_frame = ttk.LabelFrame(main_frame, text="📝 Лог", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', height=14, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        gui_handler = self._text_handler_class(self.log_text)
        gui_handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(gui_handler)

        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(main_frame, textvariable=self.status_var, relief="sunken", anchor="w", padding=(5, 2)).pack(fill=tk.X, side=tk.BOTTOM)

    def save_cookie(self):
        self.config["session_cookie"] = self.cookie_var.get().strip()
        self.save_config()
        self.logger.info("✓ Куки сохранены")
        messagebox.showinfo("Сохранено", "Куки успешно сохранены")

    def refresh_catalog_status(self):
        count = len(glob.glob(os.path.join(CATALOG_JS_DIR, "page_*.json")))
        completed = os.path.exists(CATALOG_PROGRESS_FILE)
        state = "каталог ПОЛНЫЙ ✓" if completed else "каталог неполный"
        self.catalog_status_var.set(f"Страниц в каталоге: {count} ({state})")

    def set_captions(self, captions):
        for key, text in captions.items():
            self.caption_labels[key].config(text=text)

    def update_stats(self, stats, current, total):
        self.root.after(0, self._update_stats_ui, dict(stats), current, total)

    def _update_stats_ui(self, stats, current, total):
        self.stats_labels["total"].config(text=str(stats.get("total_pages", 0)))
        found_key = "skipped" if "skipped" in stats else "found"
        self.stats_labels["found"].config(text=str(stats.get(found_key, 0)))
        self.stats_labels["downloaded"].config(text=str(stats.get("downloaded", 0)))
        self.stats_labels["errors"].config(text=str(stats.get("errors", 0)))
        self.stats_labels["page"].config(text=f"{current}/{total}")
        if total > 0:
            self.progress_var.set(current / total * 100)
            self.status_var.set(f"Прогресс: {current / total * 100:.1f}% ({current}/{total})")

    def check_cookie(self):
        cookie = self.cookie_var.get().strip()
        if not cookie:
            messagebox.showwarning("Ошибка", "Сначала вставьте и сохраните session-cookie")
            return None
        return cookie

    def lock_ui(self):
        self.btn_catalog.config(state="disabled")
        self.btn_year.config(state="disabled")
        self.btn_reset_year.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress_var.set(0)

    def unlock_ui(self, success):
        self.btn_catalog.config(state="normal")
        self.btn_year.config(state="normal")
        self.btn_reset_year.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.refresh_catalog_status()
        self.status_var.set("✓ Готово" if success else "✗ Завершено с ошибками / остановлено")

    def reset_year(self):
        year = self.year_var.get()
        year_value = "all" if year == "Все годы" else year
        base_dir = f"Реестр ОТТС_{year_value}" if year_value != "all" else "Реестр ОТТС_все_годы"
        processed_file = os.path.join(base_dir, "processed_ids.json")
        xlsx_path = os.path.join(base_dir, f"otts_{year_value}.xlsx")

        msg = f"Сбросить прогресс обработки года {year}?\n\n"
        msg += f"Будут удалены:\n• {processed_file}\n• {xlsx_path}\n\n"
        msg += f"Папка pdf/ НЕ будет тронута.\n\nПродолжить?"

        if not messagebox.askyesno("Сброс года", msg):
            return

        removed = []
        for p in (processed_file, xlsx_path):
            if os.path.exists(p):
                try:
                    os.remove(p)
                    removed.append(os.path.basename(p))
                except Exception as e:
                    self.logger.error(f"Не удалось удалить {p}: {e}")
        if removed:
            self.logger.info(f"✓ Сброшено для года {year}: {', '.join(removed)}")
            messagebox.showinfo("Готово", f"Удалено: {', '.join(removed)}\n\nЗапустите обработку заново.")
        else:
            messagebox.showinfo("Нечего сбрасывать", "Файлы прогресса не найдены.")

    def start_catalog(self):
        cookie = self.check_cookie()
        if not cookie:
            return
        self.lock_ui()
        self.set_captions({"total": "Всего страниц:", "found": "Уже есть (пропущено):",
                           "downloaded": "Скачано страниц:", "errors": "Ошибок:", "page": "Страница:"})
        self.logger.info(f"\n{'=' * 60}\nФАЗА 1: Скачивание каталога\n{'=' * 60}")
        self.worker = CatalogDownloader(cookie, self.logger, self.update_stats)
        self.thread = threading.Thread(target=self._worker_run, daemon=True)
        self.thread.start()

    def start_year(self):
        cookie = self.check_cookie()
        if not cookie:
            return
        if not glob.glob(os.path.join(CATALOG_JS_DIR, "page_*.json")):
            messagebox.showwarning("Каталог пуст", "Сначала скачайте каталог.")
            return
        year = self.year_var.get()
        year_value = "all" if year == "Все годы" else year
        self.lock_ui()
        self.set_captions({"total": "Страниц каталога:", "found": "Найдено за год:",
                           "downloaded": "Скачано PDF:", "errors": "Ошибок:", "page": "Страница:"})
        self.logger.info(f"\n{'=' * 60}\nФАЗА 2: Обработка года {year}\n{'=' * 60}")
        self.worker = YearProcessor(year_value, cookie, self.logger, self.update_stats)
        self.thread = threading.Thread(target=self._worker_run, daemon=True)
        self.thread.start()

    def _worker_run(self):
        try:
            success = self.worker.run()
        except Exception as e:
            self.logger.error(f"Неожиданная ошибка: {e}")
            success = False
        self.root.after(0, self.unlock_ui, success)

    def stop_worker(self):
        if self.worker:
            self.worker.stop_flag = True
            self.logger.info("\n⏸ Запрашиваем остановку...")
            self.btn_stop.config(state="disabled")
            self.status_var.set("Остановка...")


def main():
    root = tk.Tk()
    OTTSApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
