# -*- coding: utf-8 -*-
"""
Интеграция с Google Sheets.
Читает данные из таблицы и возвращает в том же формате что и Excel.
"""
import os
import json
import io
import pandas as pd
import numpy as np
from datetime import datetime

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]

# ID таблицы — можно задать через переменную окружения
SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID', '1nfg7-UTUTlMO3Aq3VUQpsMhDf0J7iS_Y1tacw7MA2O0')

# Листы которые читаем (можно несколько для разных периодов)
SHEET_NAMES = os.getenv('GOOGLE_SHEET_NAMES', 'Чеки ЗДР 26,Чеки увел 26').split(',')

# Кэш последней загрузки
_cache = {
    'data': None,
    'loaded_at': None,
    'spreadsheet_id': None,
}


def _get_credentials():
    """Получает credentials из переменных окружения или файла."""

    # Способ 1 — полный JSON в одной переменной GOOGLE_CREDENTIALS
    creds_json = os.getenv('GOOGLE_CREDENTIALS')
    if creds_json:
        try:
            info = json.loads(creds_json)
            return Credentials.from_service_account_info(info, scopes=SCOPES)
        except Exception as e:
            print(f'[GSheets] Ошибка GOOGLE_CREDENTIALS: {e}')

    # Способ 2 — отдельные переменные (Railway-friendly)
    private_key = os.getenv('GOOGLE_PRIVATE_KEY', '')
    # Railway иногда экранирует \n — исправляем
    private_key = private_key.replace('\\n', '\n')

    client_email = os.getenv('GOOGLE_CLIENT_EMAIL', '')
    if private_key and client_email:
        info = {
            "type": os.getenv('GOOGLE_TYPE', 'service_account'),
            "project_id": os.getenv('GOOGLE_PROJECT_ID', ''),
            "private_key_id": os.getenv('GOOGLE_PRIVATE_KEY_ID', ''),
            "private_key": private_key,
            "client_email": client_email,
            "client_id": os.getenv('GOOGLE_CLIENT_ID', ''),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
        try:
            return Credentials.from_service_account_info(info, scopes=SCOPES)
        except Exception as e:
            print(f'[GSheets] Ошибка сборки credentials: {e}')

    # Способ 3 — файл credentials.json
    for path in ['credentials.json', os.path.join(os.path.dirname(os.path.dirname(__file__)), 'credentials.json')]:
        if os.path.exists(path):
            return Credentials.from_service_account_file(path, scopes=SCOPES)

    raise FileNotFoundError(
        'Google credentials не найдены. Установите переменные окружения: '
        'GOOGLE_PRIVATE_KEY, GOOGLE_CLIENT_EMAIL, GOOGLE_PROJECT_ID'
    )


def load_from_google_sheets(spreadsheet_id: str = None, sheet_names: list = None) -> pd.DataFrame:
    """
    Загружает данные из Google Sheets и возвращает DataFrame
    в том же формате что load_and_parse из Excel.
    """
    if not GSPREAD_AVAILABLE:
        raise ImportError('Установите: pip install gspread google-auth')

    sid = spreadsheet_id or SPREADSHEET_ID
    sheets = sheet_names or SHEET_NAMES

    creds = _get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(sid)

    print(f'[GSheets] Открываем таблицу: {spreadsheet.title}')
    print(f'[GSheets] Листы: {[ws.title for ws in spreadsheet.worksheets()]}')

    dfs = []
    for sheet_name in sheets:
        sheet_name = sheet_name.strip()
        try:
            ws = spreadsheet.worksheet(sheet_name)
            print(f'[GSheets] Читаем лист "{sheet_name}"...')
            data = ws.get_all_values()
            if not data:
                print(f'[GSheets] Лист "{sheet_name}" пустой, пропускаем')
                continue
            df = pd.DataFrame(data[1:], columns=data[0])
            df['_источник'] = sheet_name
            dfs.append(df)
            print(f'[GSheets] Лист "{sheet_name}": {len(df):,} строк')
        except gspread.WorksheetNotFound:
            print(f'[GSheets] Лист "{sheet_name}" не найден, пропускаем')

    if not dfs:
        raise ValueError(f'Не найдено ни одного листа из: {sheets}')

    result = pd.concat(dfs, ignore_index=True)
    print(f'[GSheets] Итого: {len(result):,} строк')
    return result


def get_cached_or_load(spreadsheet_id: str = None, force_reload: bool = False) -> pd.DataFrame:
    """
    Возвращает кэшированные данные или загружает заново.
    Кэш живёт до следующего вызова с force_reload=True.
    """
    global _cache
    sid = spreadsheet_id or SPREADSHEET_ID

    if not force_reload and _cache['data'] is not None and _cache['spreadsheet_id'] == sid:
        age = (datetime.now() - _cache['loaded_at']).seconds
        print(f'[GSheets] Из кэша (загружено {age}с назад)')
        return _cache['data']

    raw = load_from_google_sheets(sid)
    _cache = {
        'data': raw,
        'loaded_at': datetime.now(),
        'spreadsheet_id': sid,
    }
    return raw


def is_configured() -> bool:
    """Проверяет что Google Sheets настроен."""
    if not GSPREAD_AVAILABLE:
        return False
    try:
        _get_credentials()
        return True
    except Exception:
        return False
