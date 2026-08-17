# -*- coding: utf-8 -*-
"""
Справочник тарифов ТК по регионам.
Источник: Аналитика_ТК_для_ЗДР.xlsx, лист Срок+выкуп+тариф
Курьерская своя: 100+ тыс. населения — 400₽, выезд (прочие) — 500₽
"""
import pandas as pd
import numpy as np
import re

def _clean_tarif(v):
    if pd.isna(v): return None
    s = str(v).strip().replace(',', '.').replace(' ', '')
    s = re.sub(r'[^\d.]', '', s)
    try:
        return float(s) if s else None
    except:
        return None

def _clean_region(v):
    if pd.isna(v): return None
    return str(v).strip().rstrip('.').lower()

def load_tariffs(filepath: str) -> dict:
    """
    Загружает тарифы из файла.
    Возвращает словарь: {region_key: {tk: tarif_rub}}
    """
    xl = pd.ExcelFile(filepath)
    if 'Срок+выкуп+тариф' not in xl.sheet_names:
        return {}

    raw = pd.read_excel(xl, sheet_name='Срок+выкуп+тариф', header=None)

    # Колонки: 0=Рег5Post, 4=тариф5Post | 6=РегСДЭКкур, 10=тарифСДЭКкур
    #          11=нет, 14=тарифСДЭКпвз  | 17=РегПочта, 22=тарифПочта
    COL_MAP = {
        '5Post':       (0, 4),
        'СДЭК':        (6, 10),   # СДЭК курьер
        'СДЭК ПВЗ':   (6, 14),   # СДЭК ПВЗ (тот же регион, другой тариф)
        'Почта России':(17, 22),
    }

    result = {}  # region_normalized -> {tk: tarif}

    for tk, (reg_col, tar_col) in COL_MAP.items():
        for i in range(2, len(raw)):  # с 2й строки (0=заголовок блока, 1=заголовок колонок)
            reg = _clean_region(raw.iat[i, reg_col])
            tar = _clean_tarif(raw.iat[i, tar_col])
            if not reg or not tar or tar < 50:
                continue
            if reg not in result:
                result[reg] = {}
            # Если уже есть СДЭК — не перезаписываем (оставляем курьер как основной СДЭК)
            if tk == 'СДЭК ПВЗ' and 'СДЭК' in result.get(reg, {}):
                result[reg]['СДЭК ПВЗ'] = tar
            else:
                result[reg][tk] = tar

    return result


# Статический справочник (загружается при старте)
# Нормализованные ключи регионов → тарифы
TARIFFS = {}  # заполняется при загрузке файла

# Курьерская своя — фиксированные тарифы (не зависят от региона)
COURIER_TARIFF_CITY   = 400  # для городов 100+ тыс. населения
COURIER_TARIFF_REGION = 500  # выезд в прочие нп


def normalize_region_key(region: str) -> str:
    """Нормализует название региона для поиска в справочнике."""
    if not region:
        return ''
    s = str(region).lower().strip()
    # Убираем типовые суффиксы
    for suffix in [' область', ' край', ' республика', ' респ.', ' респ',
                   'республика ', ' ао', ' автономный округ', ' автономная область',
                   ' обл.', ' обл']:
        s = s.replace(suffix, '')
    s = s.rstrip(' -.')
    return s


def get_tariff_for_region(region: str, tk: str) -> float | None:
    """
    Возвращает тариф для пары регион-ТК.
    Использует нечёткое совпадение по нормализованному ключу.
    """
    if not TARIFFS:
        return None

    key = normalize_region_key(region)
    if not key:
        return None

    # Прямое совпадение
    if key in TARIFFS and tk in TARIFFS[key]:
        return TARIFFS[key][tk]

    # Нечёткое — ищем ключ который содержится в нашем или наоборот
    for stored_key, tks in TARIFFS.items():
        if (stored_key in key or key in stored_key) and tk in tks:
            return tks[tk]

    return None
