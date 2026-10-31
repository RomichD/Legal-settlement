#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматическое обновление предельных значений ПСК для МФО в index.html.

Источник: страница Банка России «Среднерыночные значения полной стоимости
потребительского кредита (займа)» — https://www.cbr.ru/statistics/bank_sector/psk/

Скрипт:
  1. находит на странице самый свежий файл *_mfo.xlsx (по дате в имени ДДММГГГГ);
  2. скачивает его и извлекает предельные значения ПСК (графа 4 таблицы);
  3. обновляет в index.html:
     - переменную MFO_PSK_QUARTER (квартал применения значений);
     - массивы MFO_PSK_UNSECURED / MFO_PSK_SECURED;
     - справочную таблицу (tbody id="mfo-psk-tbody") и её заголовок;
  4. если изменений нет — файл остаётся нетронутым (код выхода 0).

Запуск: python3 .github/scripts/update_psk.py [путь_к_index.html]
Зависимости: requests, openpyxl
"""

import io
import re
import sys

import openpyxl
import requests

PAGE_URL = 'https://www.cbr.ru/statistics/bank_sector/psk/'
BASE_URL = 'https://www.cbr.ru'
UA = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
    'Accept-Language': 'ru-RU,ru;q=0.9',
}

# Категории строки 2.3 (займы без обеспечения) и 2.2 (с иным обеспечением):
# номер строки файла ЦБ -> (maxDays, maxAmount)
UNSECURED_MAP = [
    ('2.3.1.1', 30, 30000),
    ('2.3.1.2', 30, None),
    ('2.3.2.1', 60, 30000),
    ('2.3.2.2', 60, None),
    ('2.3.3.1', 180, 30000),
    ('2.3.3.2', 180, 100000),
    ('2.3.3.3', 180, None),
    ('2.3.4.1', 365, 100000),
    ('2.3.4.2', 365, None),
    ('2.3.5.1', 'Infinity', 100000),
    ('2.3.5.2', 'Infinity', None),
]
SECURED_MAP = [
    ('2.2.1', 365),
    ('2.2.2', 'Infinity'),
]


def http_get(url, binary=False):
    r = requests.get(url, headers=UA, timeout=90)
    r.raise_for_status()
    return r.content if binary else r.text


def find_latest_mfo_file(page_html):
    """Возвращает (url, дату публикации 'ДД.ММ.ГГГГ') самого свежего файла *_mfo.xlsx."""
    links = re.findall(r'href="(/Collection/Collection/File/\d+/(\d{8})_mfo\.xlsx)"', page_html)
    if not links:
        raise SystemExit('На странице ЦБ не найдены файлы *_mfo.xlsx')

    def ymd(date_str):
        # ДДММГГГГ -> ГГГГММДД (для корректного сравнения)
        return date_str[4:] + date_str[2:4] + date_str[:2]

    best = max(links, key=lambda x: ymd(x[1]))
    published = '{}.{}.{}'.format(best[1][:2], best[1][2:4], best[1][4:])
    return BASE_URL + best[0], published


def parse_workbook(content):
    """Извлекает из xlsx квартал и предельные значения ПСК по номерам строк."""
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb[wb.sheetnames[0]]

    limits = {}
    quarter = None
    for row in ws.iter_rows(min_col=2, max_col=5):
        num = row[0].value
        # Номер строки файла (2.3.1.1 и т.п.) -> предельное значение (графа 4)
        if isinstance(num, str):
            num = num.strip()
            if re.fullmatch(r'2\.\d+(\.\d+)*', num):
                limit = row[3].value
                if isinstance(limit, (int, float)) and limit > 0:
                    limits[num] = float(limit)
        # Квартал применения ищем в тексте ячеек второй колонки
        for cell in row[:1]:
            v = cell.value
            if isinstance(v, str) and 'квартале' in v:
                m = re.search(r'([IVX]+)\s+квартале\s+(\d{4})\s+года', v)
                if m:
                    quarter = '{} квартал {}'.format(m.group(1), m.group(2))
    if quarter is None:
        raise SystemExit('Не удалось определить квартал применения значений ПСК')

    unsecured = []
    for row_num, max_days, max_amount in UNSECURED_MAP:
        if row_num not in limits:
            raise SystemExit('В файле ЦБ не найдена строка {} (без обеспечения)'.format(row_num))
        unsecured.append({'row': row_num, 'maxDays': max_days, 'maxAmount': max_amount,
                          'psk': limits[row_num]})
    secured = []
    for row_num, max_days in SECURED_MAP:
        if row_num not in limits:
            raise SystemExit('В файле ЦБ не найдена строка {} (с обеспечением)'.format(row_num))
        secured.append({'row': row_num, 'maxDays': max_days, 'psk': limits[row_num]})
    return quarter, unsecured, secured


def js_num(v):
    """Число для JS-литерала: 292 вместо 292.0, Infinity без кавычек."""
    if v == 'Infinity':
        return 'Infinity'
    if float(v).is_integer():
        return str(int(v))
    return repr(float(v))


def js_amount(v):
    return 'null' if v is None else str(v)


def js_days(v):
    return 'Infinity' if v == 'Infinity' else str(v)


def build_array_literal(name, items):
    rows = []
    for it in items:
        amount = ('maxAmount: ' + js_amount(it['maxAmount']) + ', ') if 'maxAmount' in it else ''
        rows.append('    {{ {a}maxDays: {d}, psk: {p} }}'.format(
            a=amount, d=js_days(it['maxDays']), p=js_num(it['psk'])))
    return 'const {} = [\n{}\n];'.format(name, ',\n'.join(rows))


def ru_num(v):
    """Число для HTML-таблицы: запятая как разделитель."""
    s = '{:g}'.format(v)
    return s.replace('.', ',')


def days_label(max_days):
    return {
        30: 'до 30 дней', 60: '31–60 дней', 180: '61–180 дней',
        365: '181–365 дней', 'Infinity': 'свыше 365 дней',
    }[max_days]


def amount_label(max_amount, max_days):
    if max_amount is None:
        return 'свыше 100 000 ₽' if max_days not in (30, 60) else 'свыше 30 000 ₽'
    return 'до 100 000 ₽' if max_amount == 100000 else 'до 30 000 ₽'


def build_ref_rows(unsecured, secured):
    rows = []
    for it in unsecured:
        rows.append('<tr><td>Без обеспечения, {}, {}</td><td>{}</td></tr>'.format(
            days_label(it['maxDays']),
            amount_label(it['maxAmount'], it['maxDays']), ru_num(it['psk'])))
    for it in secured:
        rows.append('<tr><td>С иным обеспечением (кроме залога), {}</td><td>{}</td></tr>'.format(
            days_label(it['maxDays']), ru_num(it['psk'])))
    return '\n                    ' + '\n                    '.join(rows)


def main():
    index_path = sys.argv[1] if len(sys.argv) > 1 else 'index.html'

    print('Загружаю страницу статистики ПСК…')
    page_html = http_get(PAGE_URL)
    file_url, published = find_latest_mfo_file(page_html)
    print('Самый свежий файл МФО: {} (опубликован {})'.format(file_url, published))

    quarter, unsecured, secured = parse_workbook(http_get(file_url, binary=True))
    print('Квартал применения значений: {}'.format(quarter))
    for it in unsecured + secured:
        print('  строка {}: ПСК {}%'.format(it['row'], ru_num(it['psk'])))

    with open(index_path, encoding='utf-8') as f:
        html = f.read()
    original = html

    # 1) Квартал в переменной MFO_PSK_QUARTER
    html, n = re.subn(
        r"(var MFO_PSK_QUARTER = ')[^']*(';)",
        r'\g<1>' + quarter + r'\g<2>',
        html)
    if n != 1:
        raise SystemExit('Не найдена переменная MFO_PSK_QUARTER (найдено: {})'.format(n))

    # 2) Массив MFO_PSK_UNSECURED
    html, n = re.subn(
        r'const MFO_PSK_UNSECURED = \[.*?\];',
        build_array_literal('MFO_PSK_UNSECURED', unsecured).replace('MFO_PSK_UNSECURED', 'MFO_PSK_UNSECURED'),
        html, flags=re.S)
    if n != 1:
        raise SystemExit('Не найден массив MFO_PSK_UNSECURED (найдено: {})'.format(n))

    # 3) Массив MFO_PSK_SECURED
    html, n = re.subn(
        r'const MFO_PSK_SECURED = \[.*?\];',
        build_array_literal('MFO_PSK_SECURED', secured),
        html, flags=re.S)
    if n != 1:
        raise SystemExit('Не найден массив MFO_PSK_SECURED (найдено: {})'.format(n))

    # 4) Квартал в заголовке справочной таблицы и комментариях
    html, n = re.subn(
        r'(Справочно: предельные значения ПСК для МФО \(решения Банка России, )[IVX]+ квартал \d{4}( г\.\))',
        r'\g<1>' + quarter + r'\g<2>',
        html)
    if n != 1:
        print('ПРЕДУПРЕЖДЕНИЕ: заголовок справочной таблицы не обновлён (найдено: {})'.format(n))
    html, n = re.subn(
        r'(// Предельные значения ПСК для МФО — займы без обеспечения \(решения Банка России, )[IVX]+ квартал \d{4}( г\.\)\.)',
        r'\g<1>' + quarter + r'\g<2>',
        html)
    if n != 1:
        print('ПРЕДУПРЕЖДЕНИЕ: комментарий к MFO_PSK_UNSECURED не обновлён (найдено: {})'.format(n))

    # 5) Строки справочной таблицы
    html, n = re.subn(
        r'(<tbody id="mfo-psk-tbody">).*?(</tbody>)',
        r'\g<1>' + build_ref_rows(unsecured, secured) + r'\g<2>',
        html, flags=re.S)
    if n != 1:
        raise SystemExit('Не найдена таблица mfo-psk-tbody (найдено: {})'.format(n))

    if html == original:
        print('Изменений нет — данные ПСК актуальны.')
        return

    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print('index.html обновлён: значения ПСК {} (опубликованы {}).'.format(quarter, published))


if __name__ == '__main__':
    main()
