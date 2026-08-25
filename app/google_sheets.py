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
# Для нескольких таблиц используйте GOOGLE_SHEET_IDS через запятую:
# GOOGLE_SHEET_IDS=id1,id2,id3
SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID', '1nfg7-UTUTlMO3Aq3VUQpsMhDf0J7iS_Y1tacw7MA2O0')
SPREADSHEET_IDS = [sid.strip() for sid in os.getenv('GOOGLE_SHEET_IDS', SPREADSHEET_ID).split(',') if sid.strip()]

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
            # Google Sheets возвращает числа с неразрывным пробелом \xa0 как разделитель тысяч
            numeric_cols = ['Сумма', 'Стоимость доставки',
                            'Тариф за услугу по доставке и выдаче отправлений, руб. с НДС',
                            'Агентское вознаграждение, с НДС руб.',
                            'сдек Агентское вознаграждение', 'сдек Сумма за услуги',
                            'почта TARIF']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(
                        df[col].astype(str)
                            .str.replace('\xa0', '', regex=False)
                            .str.replace(' ', '', regex=False)
                            .str.replace(',', '.', regex=False),
                        errors='coerce'
                    )
            df['_источник'] = sheet_name
            dfs.append(df)
            print(f'[GSheets] Лист "{sheet_name}": {len(df):,} строк')
        except gspread.WorksheetNotFound:
            print(f'[GSheets] Лист "{sheet_name}" не найден, пропускаем')

    if not dfs:
        raise ValueError(f'Не найдено ни одного листа из: {sheets}')

    result = pd.concat(dfs, ignore_index=True)

    # ── Подтягиваем тарифы из отдельных листов (5пост, сдек, почта) ──
    def clean_num(s):
        return pd.to_numeric(
            s.astype(str).str.replace('\xa0','',regex=False)
                         .str.replace(' ','',regex=False)
                         .str.replace(',','.',regex=False),
            errors='coerce'
        )

    try:
        ws5 = spreadsheet.worksheet('5пост')
        data5 = ws5.get_all_values()
        df5 = pd.DataFrame(data5[1:], columns=data5[0])
        col5 = next((c for c in df5.columns if 'услугу по доставке' in c), None)
        key5 = next((c for c in df5.columns if 'Отправления Заказчика' in c), None)
        if col5 and key5:
            df5['_key'] = df5[key5].astype(str).str.strip()
            df5[col5] = clean_num(df5[col5])
            tarif5 = df5.groupby('_key')[col5].mean()
            result['Номер посылки'] = result['Номер посылки'].astype(str).str.strip()
            mask5 = result['Способ получения'].str.lower().str.contains('5post', na=False)
            col_dest = 'Тариф за услугу по доставке и выдаче отправлений, руб. с НДС'
            if col_dest not in result.columns:
                result[col_dest] = float('nan')
            result.loc[mask5, col_dest] = result.loc[mask5, 'Номер посылки'].map(tarif5)
            matched = result.loc[mask5, col_dest].notna().sum()
            print(f'[GSheets] 5Post тарифов найдено: {matched} из {mask5.sum()}')
        else:
            print(f'[GSheets] 5Post: col={col5}, key={key5}')
    except Exception as e:
        print(f'[GSheets] 5пост: {e}')

    try:
        wsc = spreadsheet.worksheet('сдек')
        datac = wsc.get_all_values()
        dfc = pd.DataFrame(datac[1:], columns=datac[0])
        # В Google Sheets колонка называется 'Суммазауслуги' (без пробелов)
        colc = next((c for c in dfc.columns if 'Суммазауслуги' in c or 'сумма за услуги' in c.lower()), None)
        keyc = next((c for c in dfc.columns if '№ заказа' == c.strip()), None)
        if not keyc:
            keyc = next((c for c in dfc.columns if 'заказ' in c.lower() and '№' in c), None)
        if colc and keyc:
            dfc['_key'] = dfc[keyc].astype(str).str.strip()
            dfc[colc] = clean_num(dfc[colc])
            tarifc = dfc.groupby('_key')[colc].mean()
            maskc = result['Способ получения'].str.lower().str.contains('cdek', na=False)
            if 'сдек Сумма за услуги' not in result.columns:
                result['сдек Сумма за услуги'] = float('nan')
            result.loc[maskc, 'сдек Сумма за услуги'] = \
                result.loc[maskc, 'Номер посылки'].map(tarifc)
            matched = result.loc[maskc, 'сдек Сумма за услуги'].notna().sum()
            print(f'[GSheets] СДЭК тарифов найдено: {matched} из {maskc.sum()}')
        else:
            print(f'[GSheets] СДЭК: col={colc}, key={keyc}')
    except Exception as e:
        print(f'[GSheets] сдек: {e}')

    try:
        wsp = spreadsheet.worksheet('почта')
        datap = wsp.get_all_values()
        dfp = pd.DataFrame(datap[1:], columns=datap[0])
        colp = next((c for c in dfp.columns if 'TARIF' in c or 'тариф' in c.lower()), None)
        keyp = next((c for c in dfp.columns if 'посылки' in c.lower() or 'номер' in c.lower()), None)
        if colp and keyp:
            dfp['_key'] = dfp[keyp].astype(str).str.strip()
            dfp[colp] = clean_num(dfp[colp])
            tarifp = dfp.groupby('_key')[colp].mean()
            maskp = result['Способ получения'].str.lower().str.contains('почта', na=False)
            result.loc[maskp, 'почта TARIF'] = \
                result.loc[maskp, 'Номер посылки'].map(tarifp)
            print(f'[GSheets] Почта тарифов: {maskp.sum()}')
    except Exception as e:
        print(f'[GSheets] почта: {e}')

    print(f'[GSheets] Итого: {len(result):,} строк')
    return result


def load_from_multiple_sheets(spreadsheet_ids: list = None, sheet_names: list = None) -> pd.DataFrame:
    """
    Загружает данные из нескольких Google Sheets и объединяет их.
    Используется когда данные разбиты по нескольким файлам (например по годам).
    """
    ids = spreadsheet_ids or SPREADSHEET_IDS
    all_dfs = []

    for sid in ids:
        try:
            print(f'[GSheets] Загружаем таблицу {sid[:20]}...')
            df = load_from_google_sheets(sid, sheet_names)
            df['_spreadsheet_id'] = sid
            all_dfs.append(df)
        except Exception as e:
            print(f'[GSheets] Ошибка таблицы {sid[:20]}: {e}')

    if not all_dfs:
        raise ValueError('Не удалось загрузить данные ни из одной таблицы')

    result = pd.concat(all_dfs, ignore_index=True)
    print(f'[GSheets] Всего из {len(all_dfs)} таблиц: {len(result):,} строк')
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
