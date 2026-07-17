"""Round-trip тести вірності на реальному стильованому файлі (fixture atko_extended.xlsx).

Закривають дефекти з CLAUDE.md §10 / зміни full-library-roadmap:
- D-3: зміни на завантаженій книзі через індексатор зберігаються при save
- D-4: редагування завантаженої shared-string комірки зберігається
- D-5: для файлів з 10+ листами дані не "перемішуються"
- D-6: стилі не "протікають" між комірками й книгами
- D-1 (Phase 1, поки xfail): повне читання стилів комірки зі styles.xml
"""

import os
import sys
import zipfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

import openpyxl
from openpyxl import load_workbook as oxl_load

from cytosheet import Workbook, load_workbook, Font, PatternFill, Alignment, Color

FIXTURE = os.path.join(os.path.dirname(__file__), 'fixtures', 'atko_extended.xlsx')


def test_open_existing_styled_file_basic():
    """Відкриття реального стильованого файлу не падає; значення читаються."""
    wb = load_workbook(FIXTURE)
    ws = wb.active
    assert ws.title == 'загальна таблиця-заповнити!'
    assert ws['A1'].value == 'Загальні характеристики'
    assert ws['H1'].value == 'Місцезнаходження ТКО'


def test_atko_styles_preserved_in_styles_xml_on_save(tmp_path):
    """Відкрити atko через cytosheet → зберегти → дані стилів у styles.xml збережені.

    Перевірка йде через openpyxl: він читає шрифт і заливку комірок із того самого
    styles.xml, тож якщо стилі збереглися — значення збігаються з оригіналом.
    """
    wb = load_workbook(FIXTURE)
    out = tmp_path / 'atko_passthrough.xlsx'
    wb.save(str(out))

    o = oxl_load(str(out)).active
    # H1 в оригіналі: жирний шрифт + суцільна заливка FFBDD7EE
    assert o['H1'].font.bold is True
    assert o['H1'].fill.fgColor.rgb == 'FFBDD7EE'
    assert o['H1'].value == 'Місцезнаходження ТКО'


def _read_styles_xml(path):
    """Повертає сирий вміст xl/styles.xml із xlsx (саму таблицю стилів)."""
    with zipfile.ZipFile(str(path)) as z:
        return z.read('xl/styles.xml')


def test_atko_styles_xml_table_identical_after_save(tmp_path):
    """ПРЯМЕ порівняння таблиці стилів наперед створеного файлу.

    Беремо реальний `atko_extended.xlsx` зі своїми стилями, відкриваємо нашою
    бібліотекою, зберігаємо й побайтово звіряємо `xl/styles.xml` (таблицю стилів)
    з оригіналом — вона має лишитися без втрат.
    """
    original = _read_styles_xml(FIXTURE)
    # У файлі справді є нетривіальна таблиця стилів (шрифти/заливки/cellXfs)
    assert b'<cellXfs' in original and b'<fonts' in original and b'<fills' in original

    wb = load_workbook(FIXTURE)
    out = tmp_path / 'atko_styles_pass.xlsx'
    wb.save(str(out))

    saved = _read_styles_xml(out)
    assert saved == original  # таблиця стилів збережена байт-у-байт


def test_atko_styles_xml_table_identical_after_modify(tmp_path):
    """Навіть після зміни даних таблиця `styles.xml` лишається ідентичною оригіналу."""
    original = _read_styles_xml(FIXTURE)

    wb = load_workbook(FIXTURE)
    wb.active['A1'] = 'ЗМІНЕНО'
    wb.active.cell(row=2, column=1, value=999)
    out = tmp_path / 'atko_styles_mod.xlsx'
    wb.save(str(out))

    saved = _read_styles_xml(out)
    assert saved == original


def test_cell_value_attr_marks_modified(tmp_path):
    """Scenario: Зміна через атрибут `cell.value` на завантаженій книзі зберігається.

    D-3/D-4 крізь `ws['A1'].value = x` — найпоширенішу openpyxl-ідіому. Поки
    `value` був публічним C-атрибутом, присвоєння минало `set_value`, тож лист не
    позначався зміненим, а `_shared_string_index` лишався на старому рядку.
    """
    wb = load_workbook(FIXTURE)
    ws = wb.active
    original = ws['A1'].value
    assert original and original != 'ЗМІНЕНО ЧЕРЕЗ .value'

    ws['A1'].value = 'ЗМІНЕНО ЧЕРЕЗ .value'
    assert ws._modified is True, 'присвоєння cell.value не позначило лист зміненим'
    assert ws['A1']._shared_string_index == -1, 'комірка досі посилається на старий sharedString'

    out = tmp_path / 'atko_value_attr.xlsx'
    wb.save(str(out))

    assert oxl_load(str(out)).active['A1'].value == 'ЗМІНЕНО ЧЕРЕЗ .value'


def test_load_and_save_unmodified_is_byte_identical(tmp_path):
    """Scenario: Відкрити й зберегти без змін не переписує XML листа.

    Regression-guard під ризик рішення D8: значення `value` тепер іде через
    сеттер, який позначає лист зміненим. Якби конструктор `Cell` чи парсери
    писали значення через нього, кожен завантажений лист ставав би `_modified`
    ще при розборі — і `save` перегенерував би XML усіх листів (COMPAT-2).
    """
    wb = load_workbook(FIXTURE)
    sheet_names = list(wb.sheetnames)
    assert sheet_names, 'фікстура без листів — тест нічого не доводить'

    # Сам факт завантаження (з розбором комірок) не є модифікацією
    for name in sheet_names:
        assert wb[name]._modified is False, f'лист {name} позначено зміненим при завантаженні'

    out = tmp_path / 'atko_untouched.xlsx'
    wb.save(str(out))

    with zipfile.ZipFile(FIXTURE) as zin, zipfile.ZipFile(str(out)) as zout:
        sheet_parts = [n for n in zin.namelist() if n.startswith('xl/worksheets/') and n.endswith('.xml')]
        assert sheet_parts, 'у фікстурі не знайдено XML листів'
        for part in sheet_parts:
            assert zout.read(part) == zin.read(part), f'{part} перезаписано без жодної мутації'


def test_atko_modify_persists_and_keeps_other_styles(tmp_path):
    """D-3/D-4: зміни застосовуються при save, а стилі нечіпаних комірок збережені."""
    wb = load_workbook(FIXTURE)
    ws = wb.active
    ws['A1'] = 'ЗМІНЕНО'                  # індексатор на завантаженій shared-string комірці (D-3, D-4)
    ws.cell(row=2, column=1, value=999)   # cell() на завантаженій комірці
    out = tmp_path / 'atko_mod.xlsx'
    wb.save(str(out))

    # openpyxl бачить нові значення
    o = oxl_load(str(out)).active
    assert o['A1'].value == 'ЗМІНЕНО'
    assert o['A2'].value == 999
    # і стилі нечіпаної H1 не загубилися
    assert o['H1'].font.bold is True
    assert o['H1'].fill.fgColor.rgb == 'FFBDD7EE'

    # cytosheet теж читає змінене значення назад
    wb2 = load_workbook(str(out))
    assert wb2.active['A1'].value == 'ЗМІНЕНО'
    assert wb2.active['A2'].value == 999


def test_atko_add_styled_rows_merges_into_loaded_styles(tmp_path):
    """ETL поверх існуючого файлу: дописати стильовані рядки, старі стилі цілі.

    Відкрити atko → дописати рядок 6 (A6:C6) з червоною заливкою (C6 вже існує
    в оригіналі зі своїм s=, тож перевіряється і зміна стилю існуючої комірки) →
    зберегти. Нові стилі мають влитися у styles.xml, а старі — лишитися.
    """
    wb = load_workbook(FIXTURE)
    ws = wb.active
    red = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF0000'))
    for col in (1, 2, 3):  # A6, B6, C6 (C6 існує в оригіналі)
        c = ws.cell(row=6, column=col, value=f'red-{col}')
        c.fill = red

    out = tmp_path / 'atko_added_rows.xlsx'
    wb.save(str(out))

    o = oxl_load(str(out)).active
    # Нові комірки — червоні
    for ref in ('A6', 'B6', 'C6'):
        assert o[ref].fill.patternType == 'solid', ref
        assert o[ref].fill.fgColor.rgb == 'FFFF0000', ref
    # Старі стилі нечіпаних комірок збережені
    assert o['H1'].font.bold is True
    assert o['H1'].fill.fgColor.rgb == 'FFBDD7EE'
    assert o['A1'].font.bold is True

    # cytosheet теж читає новий стиль назад
    c2 = load_workbook(str(out)).active
    assert c2['A6'].fill.fgColor.rgb == 'FFFF0000'
    assert c2['A6'].value == 'red-1'


def test_atko_change_existing_cell_style_merges(tmp_path):
    """Зміна стилю існуючої комірки зі значенням (H1) зливається, не ламаючи інші."""
    wb = load_workbook(FIXTURE)
    ws = wb.active
    ws['H1'].font = Font(name='Arial', bold=False, color=Color(rgb='FF00AA00'))

    out = tmp_path / 'atko_changed_h1.xlsx'
    wb.save(str(out))

    o = oxl_load(str(out)).active
    # H1: новий шрифт застосовано, значення збережено
    assert o['H1'].font.name == 'Arial'
    assert o['H1'].font.bold is False
    assert o['H1'].font.color.rgb == 'FF00AA00'
    assert o['H1'].value == 'Місцезнаходження ТКО'
    # H1 зберегла свою оригінальну заливку (повний Style при read → merge не губить її)
    assert o['H1'].fill.fgColor.rgb == 'FFBDD7EE'
    # Інша стильована комірка незмінна
    assert o['A1'].font.bold is True


def test_atko_empty_styled_cell_roundtrips(tmp_path):
    """Порожня комірка лише зі стилем зберігається й читається (self-closing <c s=..>)."""
    wb = load_workbook(FIXTURE)
    wb.active['A6'].fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF0000'))
    out = tmp_path / 'atko_empty_styled.xlsx'
    wb.save(str(out))

    assert oxl_load(str(out)).active['A6'].fill.fgColor.rgb == 'FFFF0000'
    assert load_workbook(str(out)).active['A6'].fill.fgColor.rgb == 'FFFF0000'


def test_cell_without_explicit_style_inherits_default():
    """Комірка без s= успадковує дефолтний стиль книги (xf=0), як openpyxl (D-1).

    Раніше такі комірки повертали порожній шрифт (name=None); тепер вони беруть
    fontId=0 зі styles.xml (а для нової книги — openpyxl-сумісний Calibri).
    """
    # Нова книга: дефолтний шрифт як в openpyxl
    wb = Workbook()
    assert wb.active['A1'].font.name == 'Calibri'
    assert wb.active['A1'].font.sz == 11.0

    # Завантажений файл, де дефолтний fontId=0 — НЕ Calibri (Arial): успадкування xf=0
    f2 = os.path.join(os.path.dirname(__file__), 'file_example_XLSX_5000.xlsx')
    if os.path.exists(f2):
        c = load_workbook(f2)['Sheet1']
        o = oxl_load(f2)['Sheet1']
        # A2 не має s= у XML, але openpyxl і cytosheet дають однаковий дефолтний шрифт
        assert c['A2'].font.name == o['A2'].font.name
        assert c['A2'].value == 1


def test_multisheet_data_mapping(tmp_path):
    """D-5: для файлу з 11 листами wb[name] повертає дані саме цього листа."""
    src = tmp_path / 'multi.xlsx'
    ob = openpyxl.Workbook()
    ob.active.title = 'First'
    ob.active['A1'] = 'First'
    for i in range(2, 12):
        s = ob.create_sheet(f'S{i}')
        s['A1'] = f'S{i}'
    ob.save(str(src))

    wb = load_workbook(str(src))
    assert wb.sheetnames == ['First'] + [f'S{i}' for i in range(2, 12)]
    for name in wb.sheetnames:
        assert wb[name]['A1'].value == name


def test_cell_style_isolation_within_and_across_workbooks():
    """D-6: зміна стилю однієї комірки не зачіпає іншу комірку чи свіжу книгу."""
    wb = Workbook()
    ws = wb.active
    ws['A1'].font = Font(bold=True)
    assert ws['B1'].font.bold is False          # сусідня комірка не зачеплена

    wb2 = Workbook()
    assert wb2.active['Z9'].font.bold is False   # свіжа книга в тому ж процесі чиста


def test_atko_cell_styles_read_matches_openpyxl():
    """D-1 (Phase 1): cell.font/fill/alignment, прочитані cytosheet, збігаються з openpyxl."""
    cyto = load_workbook(FIXTURE).active
    oxl = oxl_load(FIXTURE).active

    # H1: жирний шрифт + rgb-заливка FFBDD7EE
    assert cyto['H1'].font.bold == oxl['H1'].font.bold
    assert cyto['H1'].font.name == oxl['H1'].font.name
    assert cyto['H1'].font.sz == oxl['H1'].font.sz
    assert cyto['H1'].fill.fgColor.rgb == oxl['H1'].fill.fgColor.rgb
    assert cyto['H1'].alignment.horizontal == oxl['H1'].alignment.horizontal

    # A1: теж жирний; заливка тематична (rgb недоступний) — звіряємо лише шрифт/вирівнювання
    assert cyto['A1'].font.bold == oxl['A1'].font.bold
    assert cyto['A1'].font.name == oxl['A1'].font.name
    assert cyto['A1'].alignment.horizontal == oxl['A1'].alignment.horizontal

    # A2: нежирний (звичайна комірка з даними)
    assert cyto['A2'].font.bold == oxl['A2'].font.bold


def test_styles_xml_roundtrip_values(tmp_path):
    """D-2 (Phase 1): нова книга зі стилями -> openpyxl читає саме ці значення.

    Закриває D-2: при записі генерується повний дедуплікований styles.xml з
    реальними посиланнями cellXfs, а не лише number_format.
    """
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Styled'
    ws['A1'].font = Font(name='Arial', bold=True)
    ws['A1'].fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFFFF00'))
    ws['A1'].alignment = Alignment(horizontal='center')
    # Сусідня комірка лишається дефолтною (перевірка дедуплікації/ізоляції)
    ws['B1'] = 'plain'

    out = tmp_path / 'new_styles.xlsx'
    wb.save(str(out))

    o = oxl_load(str(out)).active
    assert o['A1'].value == 'Styled'
    assert o['A1'].font.name == 'Arial'
    assert o['A1'].font.bold is True
    assert o['A1'].fill.patternType == 'solid'
    assert o['A1'].fill.fgColor.rgb == 'FFFFFF00'
    assert o['A1'].alignment.horizontal == 'center'
    # B1 — дефолтний шрифт, без жирного
    assert o['B1'].font.bold is False
