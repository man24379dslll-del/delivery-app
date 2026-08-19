# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from data.cities_population import get_population_segment, SEGMENTS_ORDER
from data.tariffs import (load_tariffs, get_tariff_for_region,
                          COURIER_TARIFF_CITY, COURIER_TARIFF_REGION,
                          normalize_region_key)
import data.tariffs as _tariffs_module

# Глобальный кэш последнего загруженного файла (для /api/filter)
_LAST_DATA = None

def set_last_data(data):
    global _LAST_DATA
    _LAST_DATA = data

def get_last_data():
    return _LAST_DATA


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

FINAL_STATUSES = {
    'Доставлен',
    'Не доставлен | Вернулся с почты',
    'Отмена заказа',
}

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

def load_tariffs_from_xl(xl) -> dict:
    """Загружает тарифы прямо из уже открытого ExcelFile."""
    import re
    if 'Срок+выкуп+тариф' not in xl.sheet_names:
        return {}
    raw = pd.read_excel(xl, sheet_name='Срок+выкуп+тариф', header=None)

    def clean_t(v):
        if pd.isna(v): return None
        s = re.sub(r'[^\d.]', '', str(v).replace(',','.').replace(' ',''))
        try: return float(s) if s else None
        except: return None

    def clean_r(v):
        if pd.isna(v): return None
        return str(v).strip().lower().rstrip('.')

    COL_MAP = {
        '5Post':        (0, 4),
        'СДЭК':         (6, 10),
        'СДЭК ПВЗ':    (6, 14),
        'Почта России': (17, 22),
    }
    result = {}
    for tk, (rc, tc) in COL_MAP.items():
        for i in range(2, len(raw)):
            reg = clean_r(raw.iat[i, rc])
            tar = clean_t(raw.iat[i, tc])
            if not reg or not tar or tar < 50: continue
            if reg not in result: result[reg] = {}
            if tk == 'СДЭК ПВЗ':
                result[reg]['СДЭК ПВЗ'] = tar
            else:
                result[reg][tk] = tar
    return result


def load_and_parse(file_bytes: bytes) -> pd.DataFrame:
    import io
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    # Загружаем тарифный справочник из файла (если есть нужный лист)
    try:
        tariffs = load_tariffs_from_xl(xl)
        if tariffs:
            _tariffs_module.TARIFFS = tariffs
            print(f'[tariffs] Загружено {len(tariffs)} регионов из файла')
    except Exception as e:
        print(f'[tariffs skip] {e}')

    dfs = []
    for sheet in DELIVERY_SHEETS:
        if sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            df['_источник'] = sheet
            dfs.append(df)
    if not dfs:
        raise ValueError("Не найдены листы с чеками")
    data = pd.concat(dfs, ignore_index=True)

    # Исключаем "Первичный заказ"
    data = data[data['Статус'] != 'Первичный заказ'].copy()

    data['ТК']            = data['Способ получения'].apply(normalize_tk)
    data['Сегмент_города']= data['Населенный пункт'].apply(get_population_segment)
    data['Статус_группа'] = data['Статус'].apply(classify_status)
    data['Сумма']         = pd.to_numeric(data['Сумма'], errors='coerce').fillna(0)
    dc = data.get('Стоимость доставки', pd.Series(0, index=data.index))
    data['Стоимость доставки'] = pd.to_numeric(dc, errors='coerce').fillna(0)

    # ── Подтягиваем реальные тарифы из листов ТК (если есть) ──────────
    data['Номер посылки'] = data['Номер посылки'].astype(str).str.strip() \
        if 'Номер посылки' in data.columns else ''
    data['Номер посылки'] = data['Номер посылки'].replace('nan', '')
    data['тариф_факт'] = np.nan

    # 5Post
    if '5пост' in xl.sheet_names:
        try:
            df5 = pd.read_excel(xl, sheet_name='5пост')
            col5 = next((c for c in df5.columns if 'услугу по доставке' in c), None)
            if col5 and '№ Отправления Заказчика' in df5.columns:
                df5['key'] = df5['№ Отправления Заказчика'].astype(str).str.strip()
                df5[col5] = pd.to_numeric(df5[col5], errors='coerce')
                tarif5 = df5.groupby('key')[col5].mean()
                mask = data['ТК'] == '5Post'
                data.loc[mask, 'тариф_факт'] = data.loc[mask, 'Номер посылки'].map(tarif5)
        except Exception as e:
            print(f'[5Post tariff] {e}')

    # СДЭК
    if 'сдек' in xl.sheet_names:
        try:
            dfc = pd.read_excel(xl, sheet_name='сдек')
            if 'Суммазауслуги' in dfc.columns and '№ заказа' in dfc.columns:
                dfc['key'] = dfc['№ заказа'].astype(str).str.strip()
                dfc['Суммазауслуги'] = pd.to_numeric(dfc['Суммазауслуги'], errors='coerce')
                tarifc = dfc.groupby('key')['Суммазауслуги'].mean()
                mask = data['ТК'] == 'СДЭК'
                data.loc[mask, 'тариф_факт'] = data.loc[mask, 'Номер посылки'].map(tarifc)
        except Exception as e:
            print(f'[СДЭК tariff] {e}')

    # Почта
    if 'почта' in xl.sheet_names:
        try:
            dfp = pd.read_excel(xl, sheet_name='почта')
            if 'TARIF' in dfp.columns and 'Номер посылки' in dfp.columns:
                dfp['key'] = dfp['Номер посылки'].astype(str).str.strip()
                # Почта хранит тариф как строку с запятой: "1 140,50" → 1140.50
                dfp['TARIF_num'] = (dfp['TARIF'].astype(str)
                    .str.replace(' ', '', regex=False)
                    .str.replace(',', '.', regex=False))
                dfp['TARIF_num'] = pd.to_numeric(dfp['TARIF_num'], errors='coerce')
                tarifp = dfp.groupby('key')['TARIF_num'].mean()
                mask = data['ТК'] == 'Почта России'
                data.loc[mask, 'тариф_факт'] = data.loc[mask, 'Номер посылки'].map(tarifp)
        except Exception as e:
            print(f'[Почта tariff] {e}')

    # Если тариф_факт не подтянулся — пробуем справочник по региону, потом Стоимость доставки
    data['тариф_факт'] = pd.to_numeric(data['тариф_факт'], errors='coerce')
    no_tarif = data['тариф_факт'].isna()

    # Курьерская своя: 100+ тыс. = 400₽, прочие = 500₽
    mk = no_tarif & (data['ТК'] == 'Курьер (свой)')
    data.loc[mk, 'тариф_факт'] = data.loc[mk, 'Сегмент_города'].map(
        lambda s: COURIER_TARIFF_CITY if s == '100 тыс.+' else COURIER_TARIFF_REGION
    )

    # Справочник по региону для остальных ТК
    no_tarif = data['тариф_факт'].isna()
    if no_tarif.any() and _tariffs_module.TARIFFS:
        def lookup_tariff(row):
            if pd.notna(row['тариф_факт']): return row['тариф_факт']
            t = get_tariff_for_region(row.get('Регион',''), row['ТК'])
            return t if t else np.nan
        data.loc[no_tarif, 'тариф_факт'] = data[no_tarif].apply(lookup_tariff, axis=1)

    # Последний fallback — колонка Стоимость доставки
    no_tarif = data['тариф_факт'].isna()
    data.loc[no_tarif, 'тариф_факт'] = data.loc[no_tarif, 'Стоимость доставки'].replace(0, np.nan)

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


def _finrez(g):
    """Финансовый результат группы заказов (выручка - тариф туда - тариф обратно)."""
    del_rows = g[g['Статус_группа'] == 'Доставлен']
    ret_rows = g[g['Статус_группа'] == 'Возврат']
    revenue      = float(del_rows['Сумма'].sum())
    col          = 'тариф_факт' if 'тариф_факт' in g.columns else 'Стоимость доставки'

    # Средний тариф по тем строкам где он известен
    tarif_known  = g[col].dropna()
    tarif_known  = tarif_known[tarif_known > 0]
    avg_t        = float(tarif_known.mean()) if not tarif_known.empty else 0.0

    # Затраты = все заказы × средний тариф (не только те у кого тариф в данных)
    cost_fwd     = float(len(g)) * avg_t
    # Затраты на возврат = кол-во возвратов × средний тариф
    cost_ret     = float(len(ret_rows)) * avg_t

    total_cost   = cost_fwd + cost_ret
    fin          = revenue - total_cost
    roi          = round(revenue / total_cost, 1) if total_cost > 0 else None

    return {
        'revenue':    int(revenue),
        'cost_fwd':   int(cost_fwd),
        'cost_ret':   int(cost_ret),
        'total_cost': int(total_cost),
        'finrez':     int(fin),
        'roi':        roi,
        'avg_tarif':  round(avg_t, 0) if avg_t else None,
    }

def _block(g: pd.DataFrame) -> dict:
    total     = len(g)
    delivered = (g['Статус_группа'] == 'Доставлен').sum()
    returned  = (g['Статус_группа'] == 'Возврат').sum()
    cancelled = (g['Статус_группа'] == 'Отмена').sum()
    in_transit= (g['Статус_группа'] == 'В пути').sum()
    final     = delivered + returned + cancelled

    del_rows = g[g['Статус_группа'] == 'Доставлен']
    avg_chk  = del_rows['Сумма'].mean()

    # % выкупа = Дата вручения / Дата ухода с почты
    # Для Курьера (стационар/свой) — нет даты ухода, fallback на delivered/total
    col_ukhod  = 'Дата ухода с почты'
    col_vruch  = 'Дата вручения получателю'
    if col_ukhod in g.columns:
        sent     = int(g[col_ukhod].notna().sum())
        vrucheno = int((g[col_ukhod].notna() & g[col_vruch].notna()).sum()) if col_vruch in g.columns else 0
        vozvrat_sent = int((g[col_ukhod].notna() & (g['Статус_группа']=='Возврат')).sum())
        if sent > 0:
            delivery_rate = round(vrucheno / sent * 100, 1)
            return_rate   = round(vozvrat_sent / sent * 100, 1)
        else:
            # fallback для Курьера (стационар/свой)
            delivery_rate = round(delivered / total * 100, 1) if total else None
            return_rate   = round(returned  / total * 100, 1) if total else None
    else:
        sent = 0; vrucheno = 0
        delivery_rate = round(delivered / total * 100, 1) if total else None
        return_rate   = round(returned  / total * 100, 1) if total else None

    r = {
        'total':          int(total),
        'delivered':      int(delivered),
        'returned':       int(returned),
        'cancelled':      int(cancelled),
        'in_transit':     int(in_transit),
        'final':          int(final),
        'sent':           sent,
        'vrucheno':       vrucheno,
        'problematic':    int((g['Статус'] == 'Проблемный заказ').sum()) if 'Статус' in g.columns else 0,
        # % выкупа = Дата вручения / Дата ухода с почты
        'delivery_rate':  delivery_rate,
        'return_rate':    return_rate,
        'cancel_rate':    round(cancelled/total*100, 1) if total else None,
        'avg_check':      int(avg_chk) if not pd.isna(avg_chk) else 0,
        'revenue':        int(del_rows['Сумма'].sum()),
        'срок_полный_ср':      safe_mean(del_rows['срок_полный_дн']),
        'срок_полный_медиана': safe_median(del_rows['срок_полный_дн']),
        'срок_в_пути_ср':      safe_mean(del_rows['срок_в_пути_дн']),
        'срок_ожидания_ср':    safe_mean(del_rows['срок_ожидания_дн']),
    }
    r.update(_finrez(g))
    return r

def calc_summary(data):
    r = _block(data)
    r['total_revenue'] = r.pop('revenue')
    return r

def calc_by_tk(data):
    rows = []
    for tk, g in data.groupby('ТК'):
        r = _block(g); r['tk'] = tk
        r['share_pct'] = round(len(g)/len(data)*100,1)
        avg_dc = g['тариф_факт'].replace(0, np.nan).mean() \
                 if 'тариф_факт' in g.columns \
                 else g['Стоимость доставки'].replace(0, np.nan).mean()
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

MAIN_TKS = ['5Post', 'СДЭК', 'Почта России', 'Курьер (стационар)', 'Курьер (свой)']
MIN_TK_ORDERS = 10  # минимум заказов через ТК для надёжности

def calc_by_city(data, min_orders=20):
    rows = []
    for (city, region), g in data.groupby(['Населенный пункт','Регион']):
        if len(g) < min_orders: continue
        r = _block(g)
        r['city']    = city
        r['region']  = region
        r['segment'] = get_population_segment(city)

        # Разбивка по ТК внутри города
        tk_data = {}
        for tk, tg in g.groupby('ТК'):
            if tk not in MAIN_TKS: continue
            total = len(tg)
            if total < MIN_TK_ORDERS: continue
            delivered = (tg['Статус_группа']=='Доставлен').sum()
            returned  = (tg['Статус_группа']=='Возврат').sum()
            avg_dc = tg['тариф_факт'].replace(0, np.nan).mean() \
                     if 'тариф_факт' in tg.columns \
                     else tg['Стоимость доставки'].replace(0, np.nan).mean()
            days = safe_mean(tg[tg['Статус_группа']=='Доставлен']['срок_полный_дн'])
            fin  = _finrez(tg)
            tk_data[tk] = {
                'total':         int(total),
                'delivered':     int(delivered),
                'returned':      int(returned),
                'delivery_rate': round(delivered/total*100, 1) if total else None,
                'return_rate':   round(returned/total*100, 1)  if total else None,
                'avg_cost':      round(float(avg_dc), 1) if pd.notna(avg_dc) else None,
                'days':          days,
                **fin,
            }

        r['tk_breakdown'] = tk_data
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

def calc_trend_by_tk(data):
    """Тренд по месяцам отдельно для каждой ТК."""
    if 'Дата создания' not in data.columns: return {}
    d = data.dropna(subset=['Дата создания']).copy()
    d['month'] = d['Дата создания'].dt.to_period('M').astype(str)
    result = {}
    for tk, tk_group in d.groupby('ТК'):
        rows = []
        for month, g in sorted(tk_group.groupby('month')):
            r = _block(g); r['month'] = month
            rows.append(r)
        result[tk] = rows
    return result

def get_available_months(data: pd.DataFrame) -> list:
    """Возвращает список доступных месяцев в данных."""
    if 'Дата создания' not in data.columns:
        return []
    months = data['Дата создания'].dt.to_period('M').dropna().astype(str).unique()
    return sorted(months.tolist())


def filter_by_month(data: pd.DataFrame, month: str) -> pd.DataFrame:
    """Фильтрует данные по месяцу создания заказа."""
    if not month or 'Дата создания' not in data.columns:
        return data
    mask = data['Дата создания'].dt.to_period('M').astype(str) == month
    return data[mask].copy()


def calc_segment_matrix(data: pd.DataFrame) -> dict:
    """
    Матрица: для каждого сегмента города — список городов с метриками.
    Возвращает {сегмент: [{'city':..,'delivery_rate':..,'finrez':..}, ...]}
    """
    result = {}
    for seg in SEGMENTS_ORDER:
        g = data[data['Сегмент_города'] == seg]
        cities = []
        for (city, region), cg in g.groupby(['Населенный пункт','Регион']):
            if len(cg) < 10: continue
            r = _block(cg)
            cities.append({
                'city':          city,
                'region':        region,
                'total':         r['total'],
                'delivery_rate': r['delivery_rate'],
                'return_rate':   r['return_rate'],
                'finrez':        r['finrez'],
                'roi':           r['roi'],
                'avg_check':     r['avg_check'],
                'avg_tarif':     r['avg_tarif'],
            })
        cities.sort(key=lambda x: x['total'], reverse=True)
        result[seg] = cities[:50]  # топ-50 городов в каждом сегменте
    return result


def run_full_analytics(data: pd.DataFrame, month: str = '') -> dict:
    # Применяем фильтр по месяцу если задан
    filtered = filter_by_month(data, month) if month else data

    return {
        'summary':         calc_summary(filtered),
        'by_tk':           calc_by_tk(filtered),
        'by_segment':      calc_by_segment(filtered),
        'tk_by_segment':   calc_tk_by_segment(filtered),
        'by_city':         calc_by_city(filtered, min_orders=10),
        'by_region':       calc_by_region(filtered),
        'trend':           calc_trend(data),       # тренд всегда по всем данным
        'trend_by_tk':     calc_trend_by_tk(data), # тренд всегда по всем данным
        'segment_matrix':  calc_segment_matrix(filtered),
        'segments_order':  SEGMENTS_ORDER,
        'available_months': get_available_months(data),
        'active_month':    month,
    }
