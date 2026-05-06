import os
from io import BytesIO
from datetime import datetime
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "vtl_template.xlsx")

_NO_FILL = PatternFill(fill_type=None)

# 出庫 columns L-CL (col 12-90): product code → column letter
_CODE_COL_MAP = None

def _get_code_col_map():
    global _CODE_COL_MAP
    if _CODE_COL_MAP is not None:
        return _CODE_COL_MAP
    wb = openpyxl.load_workbook(TEMPLATE_PATH, read_only=True)
    ws = wb.active
    m = {}
    # Use iter_rows with explicit bounds to avoid EmptyCell/MergedCell issues
    for row_cells in ws.iter_rows(min_row=6, max_row=6, min_col=12, max_col=90):
        for cell in row_cells:
            try:
                v = cell.value
                col_num = cell.column
            except AttributeError:
                continue
            if v and isinstance(v, str):
                code = v.strip().replace(" ", "").replace("　", "").replace(" ", "")
                if len(code) >= 10:
                    m[code] = get_column_letter(col_num)
    wb.close()
    _CODE_COL_MAP = m
    return m


def build_vtl_workbook(rows):
    """
    rows: list of {
        'date':      '2026/05/06',
        'warehouse': '佰事達倉VSR-650188',
        'items':     [{'code': 'JDP0590 1A61', 'qty': 270}, ...]
    }
    Returns xlsx bytes.
    """
    code_col = _get_code_col_map()
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Set sheet tab name based on the month of the first row's date
    first_date_str = rows[0].get("date", "") if rows else ""
    try:
        first_dt = datetime.strptime(first_date_str, "%Y/%m/%d")
        ws.title = f"{first_dt.year}.{first_dt.month}月"
    except (ValueError, AttributeError):
        pass

    # Fix header row 6: reset any red/colored font to black
    for cell in ws[6]:
        try:
            if cell.font and cell.font.color and cell.font.color.type == "rgb" \
                    and cell.font.color.rgb not in ("00000000", "FF000000"):
                cell.font = cell.font.copy(color="FF000000")
        except Exception:
            pass

    def safe_set(r, c, v):
        cell = ws.cell(row=r, column=c)
        try:
            cell.value = v
            cell.fill = _NO_FILL
        except AttributeError:
            pass

    FIRST_ROW = 7
    for i, row in enumerate(rows):
        r = FIRST_ROW + i

        # Clear entire row background inherited from template
        for c in range(1, 91):
            try:
                ws.cell(row=r, column=c).fill = _NO_FILL
            except Exception:
                pass

        # Column A = 日期
        date_str = row.get("date", "")
        try:
            dt = datetime.strptime(date_str, "%Y/%m/%d")
        except ValueError:
            dt = None
        safe_set(r, 1, dt)

        # Column K = 所別/備註
        safe_set(r, 11, row.get("warehouse", ""))

        # 出庫 product quantities
        for item in row.get("items", []):
            raw = item.get("code", "")
            normalized = raw.replace(" ", "").replace("　", "").replace(" ", "")
            col_letter = code_col.get(normalized)
            if col_letter:
                cidx = column_index_from_string(col_letter)
                safe_set(r, cidx, item.get("qty", 0))

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
