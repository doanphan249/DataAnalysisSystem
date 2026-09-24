"""Tạo workbook theo cấu hình; công thức có giá trị cache để đọc ngay trong BI."""
from pathlib import Path
from datetime import date, datetime
import os
import shutil
import tempfile
import xlsxwriter
import openpyxl
from .models import DataError
from .production import METRICS, PLANT_INFO
from .sales_summary import QUANTITY_LABELS

# Power BI đọc theo tên bảng Excel. Sheet không có ở đây dùng tên bảng trùng tên sheet.
TABLE_NAMES={
 'VatTuPhuTung':'ChiPhiVTPTTable','Dien':'NangLuongDienTable',
 'TonKhoNguyenLieu':'TonKhoNguyenLieuTable','TonKhoNguyenLieuChiTiet':'ChiTietTonKhoNLTable',
 'TonKhoThanhPham':'TonKhoThanhPhamTable','TieuThuSoi':'TieuThuSoiTable',
 'TongHopKQBanHang':'TongHopKQBanHangTable','PheLieuThanhLy':'PheLieuThanhLyTable',
 'HDBong':'MuaBongTable','CongNoPhaiThu':'CongNoPhaiThuTable',
}
NOTES_SHEET = 'GhiChu'
# Quy ước mã mặt hàng trên sheet GhiChu, chép nguyên từ mẫu Dashboard.
PRODUCT_CODE_NOTES = [
    ('C3', 'theo nguyên liệu'), ('D3', 'recycle'), ('E3', 'bông+xơ tái chế (nhựa chai)'),
    ('D4', 'thường'), ('E4', 'bông+xơ nguyên sinh (dầu, mỏ)'),
    ('D5', 'cm/modal'), ('E5', 'bông+xơ gỗ'),
    ('C7', 'theo loại sợi'), ('D7', 'pha'), ('E7', 'cvc'), ('E8', 'tc'), ('E9', 'cm/modal'),
    ('C11', 'theo kỹ thuật'), ('D11', 27), ('E11', 'sợi nồi cọc ring spun'),
    ('D12', 10), ('E12', 'sợi nồi cọc compact'),
    ('C15', 'w = wax'), ('D15', 'có chuốt sáp'), ('E15', 'knitting'), ('F15', 'dệt kim'),
    ('C16', 'UW = unwax'), ('D16', 'không chuốt sáp'), ('E16', 'weaving '), ('F16', 'dệt thoi'),
    ('C17', 'wk'), ('C18', 'uww'),
]

# Định dạng hiển thị theo mẫu; không làm tròn/mất độ chính xác giá trị bên trong.
NUMBER_FORMATS = {
    'whole': '0', 'integer': '#,##0', 'decimal': '#,##0.000000', 'quantity': '#,##0.0000',
    'tons': '#,##0.000', 'flexible': '#,##0.######', 'price': '0.000000', 'price2': '0.00',
    'usd': '#,##0.00', 'date': 'yyyy-mm-dd',
}
COLUMN_FORMATS = {
    'Nam': 'whole', 'ChiSoSoi': 'whole', 'NgayHopDongPLHD': 'date',
    'GiaTri': 'integer', 'ThanhTien': 'integer', 'ThanhTienDong': 'integer',
    'GiaTriTonDong': 'integer', 'DoanhSoDong': 'integer',
    'SanLuong': 'quantity', 'SanLuongDienTieuThu': 'quantity',
    'DonGiaUSDkg': 'price', 'GiaCIFUSDkg': 'price', 'IPI': 'price', 'UPercent': 'price',
    'ThanhTienUSD': 'usd',
}
INTEGER_UNITS = {'Đồng', 'Người', 'Cọc sợi'}
PLAN_UNITS = dict(METRICS)
PLANT_UNITS = {name: unit for name, unit, _ in PLANT_INFO}


def table_name(sheet):
    return TABLE_NAMES.get(sheet, sheet)


def _this_row(sheet, column):
    return f'{table_name(sheet)}[[#This Row],[{column}]]'


def _formulas():
    """Công thức theo (sheet, cột); tham chiếu có cấu trúc nên không phụ thuộc vị trí cột."""
    m = _this_row('SanLuongSanXuat', 'MatHang')
    return {
        ('SanLuongSanXuat', 'TheoNguyenLieu'): f'=IF({m}="","",IF(ISNUMBER(SEARCH("CM/Modal",{m})),"CM/Modal",IF(ISNUMBER(SEARCH("R)",{m})),"Recycle","Thường")))',
        ('SanLuongSanXuat', 'TheoLoaiSoi'): f'=IF({m}="","",IF(ISNUMBER(SEARCH("CM/Modal",{m})),"Sợi CM/Modal",IF(ISNUMBER(SEARCH("CVC",{m})),"Sợi CVC",IF(ISNUMBER(SEARCH("TC",{m})),"Sợi TC",""))))',
        ('SanLuongSanXuat', 'LoaiChai'): f'=IF({m}="","",IF(ISNUMBER(SEARCH("CM/Modal",{m})),"Sợi chải kỹ",IF(ISNUMBER(FIND("Cm",{m})),"Sợi chải kỹ",IF(ISNUMBER(FIND("Cd",{m})),"Sợi chải thô",""))))',
        ('SanLuongSanXuat', 'ChiSoSoi'): f'=IF({m}="","",IFERROR(--LEFT({m},FIND("/",{m})-1),""))',
        ('HDBong', 'ThanhTienUSD'): f'={_this_row("HDBong", "SoLuongTan")}*1000*{_this_row("HDBong", "DonGiaUSDkg")}',
    }


FORMULAS = _formulas()
PRODUCTS = f'{table_name("SanLuongSanXuat")}[MatHang]'
PRODUCT_LIST_FORMULA = f'=SORT(UNIQUE(FILTER({PRODUCTS},{PRODUCTS}<>"")))'


def _format_for(sheet, header, row, headers):
    """Cột một giá trị nhiều đơn vị (KeHoach, GiaTri…) chọn định dạng theo đơn vị của chỉ tiêu."""
    if sheet == 'HieuQuaSanXuat' and header in ('KeHoach', 'ThucHien'):
        unit = PLAN_UNITS[row[headers.index('ChiTieu')]]
        return 'price2' if unit == 'USD/kg' else 'integer' if unit in INTEGER_UNITS else 'quantity'
    if sheet == 'ThongTinNhaMay' and header == 'GiaTri':
        return 'integer' if PLANT_UNITS.get(row[headers.index('ThongTinSanXuat')]) in INTEGER_UNITS else 'decimal'
    if sheet == 'TongHopKQBanHang' and header == 'GiaTri':
        return 'flexible' if row[headers.index('ChiTieu')] in QUANTITY_LABELS else 'integer'
    if sheet == 'HDBong' and header == 'SoLuongTan':
        return 'tons'
    if sheet == 'CongNoPhaiThu' and header != 'SanLuongBanHang':
        return 'integer'
    return COLUMN_FORMATS.get(header, 'decimal')


def _create_formats(workbook):
    base = {'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'}
    definitions = {
        'text': {},
        'header': {
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78',
            'text_wrap': True, 'align': 'center',
        },
        **{name: {'num_format': num_format} for name, num_format in NUMBER_FORMATS.items()},
    }
    return {name: workbook.add_format(base | details)
            for name, details in definitions.items()}


def product_list(report):
    """Mặt hàng khác nhau của SanLuongSanXuat; không phân biệt hoa thường như SORT/UNIQUE của Excel."""
    column = report.schemas['SanLuongSanXuat'].index('MatHang')
    return sorted({row[column] for row in report.rows['SanLuongSanXuat'] if row[column]}, key=str.casefold)


def _write_notes(workbook, formats, products):
    """Sheet GhiChu: danh sách mặt hàng (công thức mảng động kèm kết quả đã tính) và quy ước mã."""
    sheet = workbook.add_worksheet(NOTES_SHEET)
    sheet.hide_gridlines(2)
    sheet.set_column(0, 0, 32, formats['text'])
    sheet.set_column(2, 2, 40, formats['text'])
    sheet.set_column(3, 5, 28, formats['text'])
    sheet.write_string(0, 0, 'MatHangKhacNhau', formats['header'])
    sheet.write_string(0, 2, 'Quy ước mã mặt hàng (SanLuongSanXuat)', formats['header'])
    if products:
        # Excel lưu kết quả tràn ô như giá trị thường: ô đầu giữ công thức, các ô sau giữ chuỗi.
        sheet.write_dynamic_array_formula(1, 0, len(products), 0, PRODUCT_LIST_FORMULA, formats['text'], products[0])
        for index, item in enumerate(products[1:], 2):
            sheet.write_string(index, 0, item, formats['text'])
    for cell, value in PRODUCT_CODE_NOTES:
        sheet.write(cell, value, formats['text'])


def _expected_sheets(report):
    return [*report.schemas, *([NOTES_SHEET] if 'SanLuongSanXuat' in report.schemas else [])]


def write_report(report, path: Path, run_id: str) -> Path | None:
    """Ghi tạm → kiểm tra đọc lại → backup → thay file đích bằng os.replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp = tempfile.mkstemp(prefix='.report_', suffix='.xlsx', dir=path.parent)
    os.close(handle)
    temp = Path(temp)
    backup = None
    try:
        with xlsxwriter.Workbook(temp, {'strings_to_formulas': False, 'strings_to_urls': False}) as workbook:
            formats = _create_formats(workbook)
            for name, headers in report.schemas.items():
                sheet = workbook.add_worksheet(name)
                rows = report.rows[name]
                sheet.hide_gridlines(2)
                sheet.freeze_panes(1, 0)
                sheet.set_default_row(19)
                sheet.set_row(0, 34)
                for col, header in enumerate(headers):
                    content_length = max([len(header), *[len(str(r[col])) for r in rows if isinstance(r[col], str)]])
                    width = min(55, max(14, content_length + 2))
                    if any(isinstance(r[col], (int, float)) for r in rows):
                        width = max(width, 22)
                    if header in ('Nam', 'ChiSoSoi'):
                        width = max(10, len(header) + 2)
                    sheet.set_column(col, col, width, formats['text'])
                    sheet.write_string(0, col, header, formats['header'])
                # Bảng Excel cần ít nhất một dòng dữ liệu; vẫn tạo bảng khi không phát sinh để Power BI tìm thấy.
                sheet.add_table(0, 0, max(len(rows), 1), len(headers) - 1, {'name': table_name(name), 'style': 'Table Style Medium 2', 'columns': [{'header': h, 'header_format': formats['header']} for h in headers]})
                for index, row in enumerate(rows, 1):
                    for col, (header, value) in enumerate(zip(headers, row)):
                        formula = FORMULAS.get((name, header))
                        if formula:
                            fmt = formats[_format_for(name, header, row, headers)] if isinstance(value, (int, float)) else formats['text']
                            sheet.write_formula(index, col, formula, fmt, '' if value is None else value)
                        elif isinstance(value, (datetime, date)):
                            sheet.write_datetime(index, col, value, formats['date'])
                        elif isinstance(value, (float, int)):
                            sheet.write_number(index, col, value, formats[_format_for(name, header, row, headers)])
                        elif value is not None:
                            sheet.write_string(index, col, str(value), formats['text'])
            if NOTES_SHEET in _expected_sheets(report):
                _write_notes(workbook, formats, product_list(report))
            workbook.set_calc_mode('auto')
        verify_saved(temp, report)
        if path.exists():
            backup_dir = path.parent / 'backups'
            backup_dir.mkdir(exist_ok=True)
            backup = backup_dir / f'{path.stem}_{run_id}{path.suffix}'
            shutil.copy2(path, backup)
        os.replace(temp, path)
        return backup
    except PermissionError as exc:
        raise DataError(f'Không ghi được {path}. Đóng file đang mở trong Excel và kiểm tra quyền thư mục.') from exc
    finally:
        temp.unlink(missing_ok=True)


def verify_saved(path, report):
    """Đọc lại toàn bộ dữ liệu, gồm kết quả cache của các cột công thức."""
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if workbook.sheetnames != _expected_sheets(report):
            raise DataError('Sai danh sách sheet sau khi xuất')
        for name, headers in report.schemas.items():
            iterator = workbook[name].iter_rows(max_col=len(headers), values_only=True)
            if list(next(iterator)) != headers:
                raise DataError(f'Sai tiêu đề sau khi xuất {name}')
            for index, expected in enumerate(report.rows[name], 2):
                actual = list(next(iterator))
                for col, (a, e) in enumerate(zip(actual, expected), 1):
                    if isinstance(a, (float, int)) and isinstance(e, (float, int)):
                        equal = abs(a - e) <= max(1e-07, abs(e) * 1e-14)
                    elif a in (None, '') and e in (None, ''):
                        equal = True
                    else:
                        equal = a == e
                    if not equal:
                        raise DataError(f'Sai dữ liệu sau khi xuất {name}, hàng {index}, cột {col}: {a!r} != {e!r}')
        if NOTES_SHEET in workbook.sheetnames:
            saved = [r[0] for r in workbook[NOTES_SHEET].iter_rows(min_row=2, max_col=1, values_only=True) if r[0] is not None]
            if saved != product_list(report):
                raise DataError(f'Sai danh sách mặt hàng sau khi xuất {NOTES_SHEET}')
    finally:
        workbook.close()
