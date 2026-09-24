"""Tổng hợp 9 chỉ tiêu bán hàng từ phần thực hiện của sheet 1. KQ SXKD."""
from openpyxl.utils import get_column_letter

from .common import dong, key, rounded
from .models import DataError

SOURCE_SHEET = '1. KQ SXKD'
OUTPUT_SHEET = 'TongHopKQBanHang'
QUANTITY_LABELS = ('Số lượng bán nội địa', 'Số lượng bán xuất khẩu')
MONEY_LABELS = (
    'Doanh thu nội địa', 'Doanh thu xuất khẩu', 'Giá vốn hàng bán',
    'Chi phí lãi vay', 'Chi phí bán hàng',
    'Chi phí quản lý doanh nghiệp', 'Lợi nhuận trước thuế',
)


def _current_column(source, unit):
    """Phân biệt cột thực hiện với cột ước thực hiện có cùng tiêu đề con.

    Cột tiền dịch từ V sang W rồi Y trong bộ mẫu T1–T8. Đọc cả hai
    tầng tiêu đề, sau đó kiểm tra cột lũy kế liền kề và cột nội dung.
    """
    sheet = source.sheet(SOURCE_SHEET)
    matches = [
        col for col in range(4, sheet.max_column)
        if key(sheet.cell(4, col).value) == key('THỰC HIỆN THÁNG BÁO CÁO')
        and key(sheet.cell(5, col).value) == key(f'Tháng báo cáo ({unit})')
    ]
    if len(matches) != 1:
        raise DataError(f'Cần đúng một cột thực hiện ({unit}): '
                        f'{source.loc(SOURCE_SHEET, "4:5")}; thấy {matches}')
    col = matches[0]
    source.expect(SOURCE_SHEET, 5, col + 1, f'Lũy kế ({unit})')
    source.expect(SOURCE_SHEET, 5, col - 3, 'Nội dung')
    return col


def extract(source, report):
    """Đọc Tn và LKnT trực tiếp; không tự cộng lại lũy kế đã điều chỉnh.

    Tiền: triệu đồng → đồng nguyên. Lượng: tấn, làm tròn 6 chữ số.
    Ô thiếu hoặc công thức lỗi dừng chương trình, không mặc định bằng 0.
    """
    source.validate_company(SOURCE_SHEET)
    source.validate_date(SOURCE_SHEET, 'A3')
    for labels, unit, output_unit in (
        (QUANTITY_LABELS, 'tấn', 'Tấn'),
        (MONEY_LABELS, 'tr.đ', 'Đồng'),
    ):
        col = _current_column(source, unit)
        label_col = col - 3
        for label in labels:
            row = source.find(SOURCE_SHEET, label, label_col, start=6, end=15)
            # Mã đơn vị và kỳ nằm trong chính khối dữ liệu của chỉ tiêu.
            source.expect(SOURCE_SHEET, row, label_col - 4 if unit == 'tr.đ' else 3, 'SPH')
            date_col = label_col - 5 if unit == 'tr.đ' else 2
            source.validate_date(SOURCE_SHEET, f'{get_column_letter(date_col)}{row}')
            monthly = source.number(SOURCE_SHEET, row, col)
            cumulative = source.number(SOURCE_SHEET, row, col + 1)
            convert = dong if output_unit == 'Đồng' else lambda n: rounded(n, 6)
            previous = next((r[3] for r in report.rows[OUTPUT_SHEET]
                             if r[0] == source.year and r[1] == f'LK{source.month - 1}T'
                             and r[2] == label), 0)
            tolerance = 2 if output_unit == 'Đồng' else 0.000002
            if abs(previous + convert(monthly) - convert(cumulative)) > tolerance:
                report.warn('SALES_SUMMARY_CUMULATIVE_ADJUSTMENT',
                            f'{label}, T{source.month}: lũy kế nguồn khác lũy kế kỳ trước '
                            '+ tháng hiện tại; giữ nguyên lũy kế nguồn.',
                            source.loc(SOURCE_SHEET, row, col + 1))
            for period, value, source_col in (
                (f'T{source.month}', monthly, col),
                (f'LK{source.month}T', cumulative, col + 1),
            ):
                report.add(OUTPUT_SHEET, [source.year, period, label, convert(value)],
                           source.loc(SOURCE_SHEET, row, source_col))

    # Tổng lượng có cả tiêu thụ nội bộ; dùng để kiểm soát, không xuất thêm chỉ tiêu.
    col = _current_column(source, 'tấn')
    total_row = source.find(SOURCE_SHEET, 'Số lượng tiêu thụ', col - 3, 6, 15)
    detail_rows = [source.find(SOURCE_SHEET, label, col - 3, 6, 15)
                   for label in (*QUANTITY_LABELS, 'Số lượng tiêu thụ nội bộ')]
    for value_col in (col, col + 1):
        total = sum(source.number(SOURCE_SHEET, row, value_col,
                                  blank_zero=(index == 2))
                    for index, row in enumerate(detail_rows))
        report.check(f'Tổng lượng bán hàng T{source.month}, cột {value_col}', total,
                     source.number(SOURCE_SHEET, total_row, value_col),
                     source=source.loc(SOURCE_SHEET, total_row, value_col))
