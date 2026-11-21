import os
import re
import subprocess
import time
import zipfile
from pathlib import Path
from typing import List, Dict
from xml.sax.saxutils import escape

from parsers import ProductItem

_XML_ILLEGAL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def _safe_xml_text(val: object) -> str:
    return escape(_XML_ILLEGAL_RE.sub("", str(val if val is not None else "")))


HEADERS = [
    "№",
    "Площадка (Приоритет)",
    "Целевая модель (Процессор)",
    "Наличие (В продаже)",
    "Статус состояния",
    "Оценка (0-100)",
    "Выявленные нюансы / Царапины / Ремонт",
    "Товар (Название)",
    "ИТОГО С ПОШЛИНОЙ (₽)",
    "Цена витрины (₽)",
    "Таможенная пошлина (₽)",
    "Расчёт пошлины",
    "Динамика цены",
    "Продавец",
    "Тип продавца",
    "Рейтинг",
    "Отзывы",
    "Регион / Доставка (Без резерва)",
    "Ссылка на товар",
    "Описание / Характеристики",
    "Обновлено",
]

COL_WIDTHS = [6, 18, 25, 26, 30, 13, 36, 46, 20, 18, 20, 32, 18, 25, 22, 10, 10, 36, 38, 55, 18]

SHOP_PRIORITY_ORDER = {
    "Ozon": 1,
    "AliExpress": 2,
    "Яндекс Маркет": 3,
    "Авито": 4,
    "DNS": 5,
    "Wildberries": 6,
}


def _col_letter(idx_1based: int) -> str:
    res = ""
    n = idx_1based
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res


def _build_sheet_xml(items: List[ProductItem]) -> str:
    cols_xml = []
    for i, w in enumerate(COL_WIDTHS, start=1):
        cols_xml.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
    cols_block = "<cols>" + "".join(cols_xml) + "</cols>"

    rows_xml = []

    # 1. Заголовок (стиль s="1")
    header_cells = []
    for c_idx, h_text in enumerate(HEADERS, start=1):
        col_ref = f"{_col_letter(c_idx)}1"
        header_cells.append(
            f'<c r="{col_ref}" s="1" t="inlineStr"><is><t>{escape(h_text)}</t></is></c>'
        )
    rows_xml.append(f'<row r="1" ht="26" customHeight="1">{"".join(header_cells)}</row>')

    # 2. Строки данных
    for r_offset, item in enumerate(items, start=2):
        # Выбор цветового стиля строки по коду состояния:
        # s=2 (Зелёный - Годный), s=3 (Жёлтый - Под вопросом), s=4 (Красный - Брак)
        if item.status_code == "GREEN":
            row_style = 2
            price_style = 5
            link_style = 8
        elif item.status_code == "YELLOW":
            row_style = 3
            price_style = 6
            link_style = 9
        else:
            row_style = 4
            price_style = 7
            link_style = 10

        values = [
            ("n", str(r_offset - 1), row_style),
            ("s", item.shop, row_style),
            ("s", item.query, row_style),
            ("s", item.in_stock, row_style),
            ("s", item.status_label, row_style),
            ("n", str(item.score), row_style),
            ("s", item.defects_summary, row_style),
            ("s", item.title, row_style),
            ("sum", str(int(item.base_price or 0) + int(item.customs_duty or 0)), price_style),
            ("n", str(item.base_price), price_style),
            ("n", str(item.customs_duty), price_style),
            ("s", item.duty_note, row_style),
            ("s", item.price_change, row_style),
            ("s", item.seller_name, row_style),
            ("s", item.seller_type, row_style),
            ("n", str(item.seller_rating), row_style),
            ("n", str(item.reviews_count), row_style),
            ("s", item.location_delivery, row_style),
            ("f", item.url, link_style),
            ("s", item.description, row_style),
            ("s", item.updated_at, row_style),
        ]

        cells = []
        for c_idx, (v_type, val, st) in enumerate(values, start=1):
            col_ref = f"{_col_letter(c_idx)}{r_offset}"
            if v_type == "n":
                cells.append(f'<c r="{col_ref}" s="{st}"><v>{_safe_xml_text(val)}</v></c>')
            elif v_type == "sum":
                cells.append(
                    f'<c r="{col_ref}" s="{st}"><f>J{r_offset}+K{r_offset}</f><v>{_safe_xml_text(val)}</v></c>'
                )
            elif v_type == "f":
                safe_url = _safe_xml_text(str(val or "").replace('"', "%22"))
                link_label = (
                    "Открыть на AliExpress (🔥 есть дешёвая скидка!) ↗"
                    if item.shop == "AliExpress"
                    else "Открыть товар ↗"
                )
                formula = f'HYPERLINK("{safe_url}","{link_label}")'
                cells.append(
                    f'<c r="{col_ref}" s="{st}" t="str"><f>{formula}</f><v>{link_label}</v></c>'
                )
            else:
                clean_str = _safe_xml_text(val)
                cells.append(
                    f'<c r="{col_ref}" s="{st}" t="inlineStr"><is><t>{clean_str}</t></is></c>'
                )

        rows_xml.append(f'<row r="{r_offset}" ht="21" customHeight="1">{"".join(cells)}</row>')

    last_row = max(1, len(items) + 1)
    last_col = _col_letter(len(HEADERS))

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  {cols_block}
  <sheetData>
    {"".join(rows_xml)}
  </sheetData>
  <autoFilter ref="A1:{last_col}{last_row}"/>
</worksheet>"""


def _build_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0 &quot;руб. (₽)&quot;"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="10"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF111111"/><name val="Segoe UI"/></font>
    <font><u/><sz val="10"/><color rgb="FF0563C1"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="11">
    <!-- 0: Default -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <!-- 1: Header -->
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 2: GREEN row -->
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <!-- 3: YELLOW row -->
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <!-- 4: RED row -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <!-- 5: GREEN price -->
    <xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 6: YELLOW price -->
    <xf numFmtId="164" fontId="2" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 7: RED price -->
    <xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 8: GREEN link -->
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 9: YELLOW link -->
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 10: RED link -->
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"""


def save_to_excel(
    all_items: List[ProductItem],
    target_path: Path,
    city_name: str = "Все регионы РФ",
) -> Path:
    """
    Сохраняет данные в многостраничный .xlsx файл с автофильтрами и цветовой разметкой состояния.
    Порядок приоритета: 1. Ozon -> 2. AliExpress -> 3. Яндекс Маркет -> 4. Авито -> 5. DNS -> 6. Wildberries.
    """
    def _priority_sort_key(item: ProductItem):
        shop_rank = SHOP_PRIORITY_ORDER.get(item.shop, 99)
        stock_rank = 0 if item.in_stock.startswith("✅") else 1
        status_rank = 0 if item.status_code == "GREEN" else (1 if item.status_code == "YELLOW" else 2)
        return (status_rank, stock_rank, shop_rank, item.price)

    def _shop_item_key(item: ProductItem):
        stock_rank = 0 if item.in_stock.startswith("✅") else 1
        status_rank = 0 if item.status_code == "GREEN" else (1 if item.status_code == "YELLOW" else 2)
        return (status_rank, stock_rank, item.price)

    # В Excel выводим абсолютно все товары от зелёного до красного (с галочками ✅ и крестиками ❌)
    valid_items = list(all_items)

    good_by_priority = sorted(
        valid_items,
        key=_priority_sort_key,
    )
    ozon_items = sorted([i for i in valid_items if i.shop == "Ozon"], key=_shop_item_key)
    avito_items = sorted([i for i in valid_items if i.shop == "Авито"], key=_shop_item_key)
    dns_items = sorted([i for i in valid_items if i.shop == "DNS"], key=_shop_item_key)
    wb_items = sorted([i for i in valid_items if i.shop == "Wildberries"], key=_shop_item_key)
    ali_items = sorted([i for i in valid_items if i.shop == "AliExpress"], key=_shop_item_key)
    ym_items = sorted([i for i in valid_items if i.shop == "Яндекс Маркет"], key=_shop_item_key)

    short_city = (city_name or "Все регионы РФ").split("-")[0][:14]
    sheets: List[tuple[str, List[ProductItem]]] = [
        ("Сводная (По приоритету)", good_by_priority),
        ("1_Ozon", ozon_items),
        ("2_AliExpress", ali_items),
        ("3_Яндекс Маркет", ym_items),
        (f"4_Авито ({short_city}+Доставка)"[:31], avito_items),
        (f"5_DNS ({short_city})"[:31], dns_items),
        ("6_Wildberries", wb_items),
    ]

    def _write_xlsx(file_path: Path) -> None:
        with zipfile.ZipFile(file_path, "w", zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            sheet_overrides = "".join(
                f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for idx in range(1, len(sheets) + 1)
            )
            content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  {sheet_overrides}
</Types>"""
            zf.writestr("[Content_Types].xml", content_types)

            # _rels/.rels
            root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
            zf.writestr("_rels/.rels", root_rels)

            # xl/workbook.xml
            sheets_entries = "".join(
                f'<sheet name="{escape(title)}" sheetId="{idx}" r:id="rId{idx}"/>'
                for idx, (title, _) in enumerate(sheets, start=1)
            )
            workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    {sheets_entries}
  </sheets>
</workbook>"""
            zf.writestr("xl/workbook.xml", workbook_xml)

            # xl/_rels/workbook.xml.rels
            wb_rels_entries = "".join(
                f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
                for idx in range(1, len(sheets) + 1)
            )
            styles_rid = len(sheets) + 1
            wb_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {wb_rels_entries}
  <Relationship Id="rId{styles_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
            zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)

            # xl/styles.xml
            zf.writestr("xl/styles.xml", _build_styles_xml())

            # xl/worksheets/sheet{idx}.xml
            for idx, (_, sheet_items) in enumerate(sheets, start=1):
                zf.writestr(f"xl/worksheets/sheet{idx}.xml", _build_sheet_xml(sheet_items))

    try:
        _write_xlsx(target_path)
        return target_path
    except PermissionError:
        # Если книга Мониторинг_Товаров.xlsx открыта в Excel, мягко закрываем её через COM,
        # перезаписываем этот же файл (не создавая лишних файлов *_LIVE.xlsx) и открываем снова
        ps_close = (
            "try { "
            "$ex = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); "
            f"foreach ($wb in $ex.Workbooks) {{ if ($wb.Name -eq '{target_path.name}') {{ $wb.Close($false) }} }} "
            "} catch {}"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_close],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(0.8)
        _write_xlsx(target_path)
        try:
            os.startfile(str(target_path))
        except Exception:
            pass
        return target_path
