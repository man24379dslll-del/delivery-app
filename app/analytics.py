# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from data.cities_population import get_population_segment, SEGMENTS_ORDER

DELIVERY_SHEETS = ['Чеки ЗДР 26', 'Чеки увел 26']
STATUS_DELIVERED = 'Доставлен'
STATUS_RETURNED  = 'Не доставлен | Вернулся с почты'
STATUS_CANCELLED = 'Отмена заказа'

TK_MAP = {
    '5post (до пункта самовывоза)': '5Post',
    'cdek (до двери)':              'СДЭК',
    'cdek (до пункта самовывоза)':  'СДЭК',
    'почта (до пункта самовывоза)': 'Почта России',
    'почта (до двери)':             'Почта России',
    'курьерская стационарная (до двери)': 'Курьер (стационар)',
    'курьерская своя (до двери)':   'Курьер (свой)',
    'другой способ (до пункта самовывоза)': 'Другое',
}
DATE_COLS = ['Дата создания','Дата сдачи на отправку','Дата ухода с почты',
             'Дата прихода','Дата вручения получателю',
             'Дата отправки возврата','Дата вручения возврата']

def normalize_tk(v):
    if pd.isna(v): return 'Неизвестно'
    return TK_MAP.get(str(v).lower().strip(), str(v))

def classify_status(s):
    if pd.isna(s): return 'В пути'
    s = str(s).strip()
    if s == STATUS_DELIVERED: return 'Доставлен'
    if s == STATUS_RETURNED:  return 'Возврат'
    if s == STATUS_CANCELLED: return 'Отмена'
    return 'В пути'

def safe_mean(series, lo=0, hi=120):
    s = pd.to_numeric(series, errors='coerce')
    s = s[s.between(lo, hi)]
    return round(float(s.mean()), 1) if not s.empty else None

def safe_median(series, lo=0, hi=120):
    s = pd.to_numeric(series, errors='coerce')
    s = s[s.between(lo, hi)]
    return round(float(s.median()), 1) if not s.empty else None

def load_and_parse(file_bytes: bytes) -> pd.DataFrame:
    import io
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    dfs = []
    for sheet in DELIVERY_SHEETS:
        if sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            df['_источник'] = sheet
            dfs.append(df)
    if not dfs:
        raise ValueError("Не найдены листы с чеками")
    data = pd.concat(dfs, ignore_index=True)

    data['ТК']            = data['Способ получения'].apply(normalize_tk)
    data['Сегмент_города']= data['Населенный пункт'].apply(get_population_segment)
    data['Статус_группа'] = data['Статус'].apply(classify_status)
    data['Сумма']         = pd.to_numeric(data['Сумма'], errors='coerce').fillna(0)
    dc = data.get('Стоимость доставки', pd.Series(0, index=data.index))
    data['Стоимость доставки'] = pd.to_numeric(dc, errors='coerce').fillna(0)

    for col in DATE_COLS:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], dayfirst=True, errors='coerce')

    # Сроки
    has_full = data['Дата создания'].notna() & data['Дата вручения получателю'].notna()
    data['срок_полный_дн'] = np.where(
        has_full,
        (data['Дата вручения получателю'] - data['Дата создания']).dt.days,
        np.nan)

    has_tr = data['Дата ухода с почты'].notna() & data['Дата прихода'].notna()
    data['срок_в_пути_дн'] = np.where(
        has_tr,
        (data['Дата прихода'] - data['Дата ухода с почты']).dt.days,
        np.nan)

    has_pvz = data['Дата прихода'].notna() & data['Дата вручения получателю'].notna()
    data['срок_ожидания_дн'] = np.where(
        has_pvz,
        (data['Дата вручения получателю'] - data['Дата прихода']).dt.days,
        np.nan)

    return data

def _block(g: pd.DataFrame) -> dict:
    total     = len(g)
    delivered = (g['Статус_группа'] == 'Доставлен').sum()
    returned  = (g['Статус_группа'] == 'Возврат').sum()
    cancelled = (g['Статус_группа'] == 'Отмена').sum()
    del_rows  = g[g['Статус_группа'] == 'Доставлен']
    avg_chk   = del_rows['Сумма'].mean()
    return {
        'total':          int(total),
        'delivered':      int(delivered),
        'returned':       int(returned),
        'cancelled':      int(cancelled),
        'delivery_rate':  round(delivered/total*100,1) if total else None,
        'return_rate':    round(returned/total*100,1)  if total else None,
        'avg_check':      int(avg_chk) if not pd.isna(avg_chk) else 0,
        'revenue':        int(del_rows['Сумма'].sum()),
        'срок_полный_ср':      safe_mean(del_rows['срок_полный_дн']),
        'срок_полный_медиана': safe_median(del_rows['срок_полный_дн']),
        'срок_в_пути_ср':      safe_mean(del_rows['срок_в_пути_дн']),
        'срок_ожидания_ср':    safe_mean(del_rows['срок_ожидания_дн']),
    }

def calc_summary(data):
    r = _block(data)
    r['total_revenue'] = r.pop('revenue')
    return r

def calc_by_tk(data):
    rows = []
    for tk, g in data.groupby('ТК'):
        r = _block(g); r['tk'] = tk
        r['share_pct'] = round(len(g)/len(data)*100,1)
        avg_dc = g['Стоимость доставки'].replace(0, np.nan).mean()
        r['avg_delivery_cost'] = round(float(avg_dc),2) if not pd.isna(avg_dc) else None
        rows.append(r)
    return sorted(rows, key=lambda x: x['total'], reverse=True)

def calc_by_segment(data):
    rows = []
    for seg in SEGMENTS_ORDER:
        g = data[data['Сегмент_города']==seg]
        if g.empty: continue
        r = _block(g); r['segment'] = seg
        r['share_pct'] = round(len(g)/len(data)*100,1)
        rows.append(r)
    return rows

def calc_tk_by_segment(data):
    result = {}
    for tk, tg in data.groupby('ТК'):
        seg_rows = []
        for seg in SEGMENTS_ORDER:
            g = tg[tg['Сегмент_города']==seg]
            if g.empty:
                seg_rows.append({'segment':seg,'total':0,'delivery_rate':None,
                                 'срок_полный_ср':None,'срок_полный_медиана':None})
            else:
                r = _block(g); r['segment'] = seg
                seg_rows.append(r)
        result[tk] = seg_rows
    return result

def calc_by_city(data, min_orders=20):
    rows = []
    for (city, region), g in data.groupby(['Населенный пункт','Регион']):
        if len(g) < min_orders: continue
        r = _block(g)
        r['city']    = city
        r['region']  = region
        r['segment'] = get_population_segment(city)
        rows.append(r)
    rows.sort(key=lambda x: x['total'], reverse=True)
    return rows[:300]

def calc_by_region(data):
    rows = []
    for region, g in data.groupby('Регион'):
        r = _block(g); r['region'] = region
        r['share_pct'] = round(len(g)/len(data)*100,1)
        rows.append(r)
    return sorted(rows, key=lambda x: x['total'], reverse=True)

def calc_trend(data):
    if 'Дата создания' not in data.columns: return []
    d = data.dropna(subset=['Дата создания']).copy()
    d['month'] = d['Дата создания'].dt.to_period('M').astype(str)
    rows = []
    for month, g in sorted(d.groupby('month')):
        r = _block(g); r['month'] = month
        rows.append(r)
    return rows

def run_full_analytics(data: pd.DataFrame) -> dict:
    return {
        'summary':        calc_summary(data),
        'by_tk':          calc_by_tk(data),
        'by_segment':     calc_by_segment(data),
        'tk_by_segment':  calc_tk_by_segment(data),
        'by_city':        calc_by_city(data, min_orders=20),
        'by_region':      calc_by_region(data),
        'trend':          calc_trend(data),
        'segments_order': SEGMENTS_ORDER,
    }
