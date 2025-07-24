import os
import sys
import pandas as pd
import re
from docx import Document
from datetime import datetime, timedelta
from pymorphy2 import MorphAnalyzer

from manual_declensions_ua import manual_declensions_ua

# Морфологический анализ для украинских имен
DICT_PATH = os.path.join(os.path.dirname(__file__), 'pymorphy2_dicts_uk/data')
morph = MorphAnalyzer(lang='uk', path=DICT_PATH)

# Функція для коректного визначення шляхів до ресурсів при запуску з exe
def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        # Запущено з .exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Запущено з .py
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

EXCEL_FILE       = resource_path('Дані подача РАПОРТІВ.xlsx')
OUTPUT_DIR       = resource_path('Результати')

TEMPLATE_SWITCH = {
    '156': {
        'dovidka': resource_path('ШАБЛОНИ/доповідь - Шаблон.docx'),
        'raport': resource_path('ШАБЛОНИ/РАПОРТ СЗЧ - ШАБЛОН.docx'),
    },
    '71': {
        'raport': resource_path('ШАБЛОНИ/РАПОРТ СЗЧ - ШАБЛОН 71а.docx'),
    },
}

VAR_TO_COL = {
    'ДАННІ_ПРИЗВ':      'ДАННІ_ПРИЗВ',
    'ДАННІ_ІНН':        'ДАННІ_ІНН',
    'ДАННІ_НР':         'ДАННІ_НР',
    'ДАННІ_МПРОЖИВ':    'ДАННІ_МПРОЖИВ',
    'ДАННІ_ТЕЛ':        'ДАННІ_ТЕЛ',
    'ДАННІ_ОСВ':        'ДАННІ_ОСВ',
    'ДАННІ_ВІРА':       'ДАННІ_ВІРА',
    'ДАННІ_РОДИЧ':      'ДАННІ_РОДИЧ',
    'ДАННІ_СІМЕЙСТАН':  'ДАННІ_СІМЕЙСТАН',
    'ДАННІ_НАЦІОН':     'ДАННІ_НАЦІОН',
    'ДАННІ_ВІЙСЬКОВИЙ': 'ДАННІ_ВІЙСЬКОВИЙ',
    'ДАННІ_ПАСПОРТ':    'ДАННІ_ПАСПОРТ',
}

MONTHS = ['січня','лютого','березня','квітня','травня','червня',
          'липня','серпня','вересня','жовтня','листопада','грудня']

def get_investigation_deadline(date, days=10):
    # Парсимо дату у форматі "дд.мм.рррр"
    #base_date = datetime.strptime(date_str, "%d.%m.%Y")
    deadline_date = date + timedelta(days=days)
    return deadline_date.strftime("%d.%m.%Y")

def decline_text(text, case):
    text_lower = text.lower().strip()
    if text_lower in manual_declensions_ua and case in manual_declensions_ua[text_lower]:
        return manual_declensions_ua[text_lower][case]
    declined = []
    for word in text.split():
        parsed = morph.parse(word)[0]
        inflected = parsed.inflect({case})
        declined.append(inflected.word if inflected else word)
    return ' '.join(declined)

def extract_department_fields(title):
    # шукаємо все, що йде після "заступник командира"
    match = list(re.finditer(r"\d+\s+\S+\s+роти", title, flags=re.IGNORECASE))
    if match:
        # Починаємо з позиції останньої роти
        start = match[-1].start()
        department_nom = title[start:].strip()
    else:
        # Фолбек: усе після останнього дефісу або увесь рядок
        parts = re.split(r"\s*[—–\-]\s*", title)
        department_nom = parts[-1].strip() if parts else title.strip()
    department_gen = decline_text(department_nom, "gent")
    return department_nom, department_gen

def date_parts(dt):
    return {
        'СЗЧ_ДАТА':      dt.strftime('%d.%m.%Y'),
        'ДОПОВІДЬ_ДАТА': dt.strftime('%d.%m.%Y'),
        'ДОПОВІДЬ_ДЕНЬ': str(dt.day),
        'ДОПОВІДЬ_МІСЯЦЬ':MONTHS[dt.month-1],
        'ДОПОВІДЬ_РІК':  str(dt.year),
        'РАПОРТ_ДАТА':   dt.strftime('%d.%m.%Y'),
        'РАПОРТ_ДЕНЬ':   str(dt.day),
        'РАПОРТ_МІСЯЦЬ': MONTHS[dt.month-1],
        'РАПОРТ_РІК':    str(dt.year),
    }

# ——— Helper-функції для форматування ———
def format_date(val):
    """Повертає дату у форматі 'дд.мм.рррр' або пустий рядок."""
    if pd.notna(val):
        return pd.to_datetime(val).strftime('%d.%m.%Y')
    return ''

def format_phone(val):
    """Повертає телефон у форматі '(xxx) xxx-xx-xx' або оригінальне значення."""
    if pd.isna(val):
        return ''
    digits = re.sub(r'\D', '', str(val))
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:8]}-{digits[8:10]}"
    return str(val)

def format_inn(val):
    """Повертає ІПН як ціле число без .0 або оригінальне значення."""
    if pd.isna(val):
        return ''
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val)

def default_format(val):
    """Загальний форматер: рядок, якщо не NaN і не 'N/A'."""
    if pd.notna(val) and str(val).upper() != 'N/A':
        return str(val)
    return ''

# ——— Словник зв’язків назва поля → функція форматування ———

formatters = {
    'ДАННІ_НР':  format_date,
    'ДАННІ_ТЕЛ': format_phone,
    'ДАННІ_ІНН': format_inn,
}

def to_genitive(text):
    if not isinstance(text, str): return ''
    words = text.split()
    inflected = [morph.parse(w)[0].inflect({'gent'}).word if morph.parse(w)[0].inflect({'gent'}) else w for w in words]
    return ' '.join(inflected)

def format_name(text):
    if not isinstance(text, str): return ''
    parts = text.split()
    if not parts: return ''
    surname = parts[0].upper()
    rest = [w.capitalize() for w in parts[1:]]
    return ' '.join([surname] + rest)

def replace_placeholders(doc, ctx):
    for p in doc.paragraphs:
        for key, val in ctx.items():
            if f'{{{{{key}}}}}' in p.text and val is not None:
                for run in p.runs:
                    run.text = run.text.replace(f'{{{{{key}}}}}', val)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                replace_placeholders(cell, ctx)

def short_fio(full_name):
    if not full_name or not isinstance(full_name, str): return ""
    parts = full_name.strip().split()
    if len(parts) < 2: return full_name
    surname = parts[0]
    name_initial = parts[1][0] + '.' if len(parts[1]) > 0 else ''
    return f"{name_initial} {surname}"

os.makedirs(OUTPUT_DIR, exist_ok=True)
df = pd.read_excel(EXCEL_FILE)

for _, row in df.iterrows():
    format_value = str(row.get('ФОРМАТ', '')).strip()
    templates = TEMPLATE_SWITCH.get(format_value, {})

    if not templates:
        print(f"[УВАГА] Формат '{format_value}' не знайдено у TEMPLATE_SWITCH — пропущено.")
        continue

    fio = row.get('ФИО', '')
    if not isinstance(fio, str) or not fio.strip():
        continue

    ctx = {}
    unit = str(row.get('ВІЙСЬК_ЧАСТИНА', '') or '')

    # ПОСАДА
    base_posada = str(row.get('ПОСАДА', '') or '')
    department, _ = extract_department_fields(base_posada)
    ctx['ПОСАДА'] = base_posada
    ctx['ПОСАДА_РОД'] = to_genitive(base_posada)

    # СЛІДЧИЙ
    ctx['СЛІДЧИЙ_ПОСАДА'] = row.get('СЛІДЧИЙ_ПОСАДА', '')
    ctx['СЛІДЧИЙ_ПОСАДА_РОД'] = to_genitive(ctx['СЛІДЧИЙ_ПОСАДА'])
    ctx['СЛІДЧИЙ_ЗВАННЯ'] = str(row.get('СЛІДЧИЙ_ЗВАННЯ','') or '')
    ctx['СЛІДЧИЙ_ФИО'] = format_name(str(row.get('СЛІДЧИЙ_ФИО','') or ''))
    ctx['СЛІДЧИЙ_ЗВАННЯ_РОД'] = to_genitive(ctx['СЛІДЧИЙ_ЗВАННЯ'])
    ctx['СЛІДЧИЙ_ФИО_РОД'] = format_name(to_genitive(ctx['СЛІДЧИЙ_ФИО']))

    # ПОДАВАЧ
    ctx['ПОДАВАЧ_ПОСАДА'] = str(row.get('ПОДАВАЧ_ПОСАДА', '') or '')
    ctx['ПОДАВАЧ_ПОСАДА_РОД'] = to_genitive(ctx['ПОДАВАЧ_ПОСАДА'])
    ctx['ПОДАВАЧ_ЗВАННЯ'] = str(row.get('ПОДАВАЧ_ЗВАННЯ','') or '')
    ctx['ПОДАВАЧ_ФИО'] = str(row.get('ПОДАВАЧ_ФИО','') or '')
    ctx['ПОДАВАЧ_ЗВАННЯ_РОД'] = to_genitive(ctx['ПОДАВАЧ_ЗВАННЯ'])
    ctx['ПОДАВАЧ_ФИО_РОД'] = format_name(to_genitive(ctx['ПОДАВАЧ_ФИО']))
    ctx['ПОДАВАЧ_ФИО_КРАТКО'] = short_fio(ctx['ПОДАВАЧ_ФИО'])

    ctx['ФИО'] = format_name(fio)
    ctx['ЗВАННЯ'] = str(row.get('ЗВАННЯ','') or '')
    ctx['ЗВАННЯ_РОД'] = to_genitive(ctx['ЗВАННЯ'])
    ctx['ФИО_РОД'] = format_name(to_genitive(fio))

    ctx['ТЕРРІТОРІЯ_ЦПП'] = str(row.get('ТЕРРІТОРІЯ_ЦПП','') or '')
    ctx['ТЕРРІТОРІЯ_НП'] = str(row.get('ТЕРРІТОРІЯ_НП','') or '')
    ctx['ТЕРРІТОРІЯ_РН'] = str(row.get('ТЕРРІТОРІЯ_РН','') or '')
    ctx['ДОПОВІДЬ_ВІД'] = str(row.get('ДОПОВІДЬ_ВІД','') or '')
    ctx['ПІДРОЗДІЛ'] = department
    ctx['ЗБРОЯ'] = str(row.get('ЗБРОЯ','') or '')

    tval = row.get('ЧАС_ДОП_ЧЕРГ')
    tstr = str(tval) if pd.notna(tval) else ''
    ctx['ЧАС_ДОП_ЧЕРГ'] = tstr.split(':')[0] + ':' + tstr.split(':')[1] if ':' in tstr else tstr

    date_val = row.get('СЗЧ_ДАТА')
    if pd.notna(date_val):
        dt = pd.to_datetime(date_val)
        ctx.update(date_parts(dt))
    else:
        ctx['СЗЧ_ДАТА'] = ''

    tval = row.get('СЗЧ_ЧАС')
    ctx['ТЕРМІН_РОЗСЛІДУВАННЯ'] = get_investigation_deadline(date_val, 11)
    tstr = str(tval) if pd.notna(tval) else ''
    ctx['СЗЧ_ЧАС'] = tstr.split(':')[0] + ':' + tstr.split(':')[1] if ':' in tstr else tstr

    obst = row.get('ОБСТАВИНИ')
    ctx['ОБСТАВИНИ'] = ', ' + str(obst).strip() if pd.notna(obst) and str(obst).strip() else ''

    for var, col in VAR_TO_COL.items():
        raw = row.get(col)
        formatter = formatters.get(var, default_format)
        ctx[var] = formatter(raw)

    ctx['ДОПОВІДЬ_НОМЕР'] = ''

    suffix = fio.split()[0].upper() + ' ' + ''.join([p[0].upper()+'.' for p in fio.split()[1:]])

    # Генерація довідки
    if 'dovidka' in templates:
        doc = Document(templates['dovidka'])
        replace_placeholders(doc, ctx)
        filename = f"А5003 Довідка_{suffix}.docx"
        doc.save(os.path.join(OUTPUT_DIR, filename))

    # Генерація рапорту
    if 'raport' in templates:
        doc = Document(templates['raport'])
        replace_placeholders(doc, ctx)
        szch_date = ctx.get('СЗЧ_ДАТА', '').strip()
        filename = f"А5003 РАПОРТ {fio.split()[0].upper()} {szch_date}.docx"
        doc.save(os.path.join(OUTPUT_DIR, filename))

print('Готово! Документи збережено в', OUTPUT_DIR)
