"""Công nợ phải thu theo nước/kênh từ sheet 2. Phai thu; tiền đổi từ triệu đồng sang đồng."""
from .common import country_name, dong, key, rounded, text
from .models import DataError

SOURCE_SHEET = '2. Phai thu'
OUTPUT_SHEET = 'CongNoPhaiThu'
# Cột đầu ra → tiêu đề nguồn. Doanh số lũy kế có năm ghi sai (LK 2025/2026)
# giữa các tháng nên chỉ so khớp phần đầu.
MAIN_HEADERS = {
    'SanLuongBanHang': 'Sản lượng bán hàng',
    'DoanhSoThangDong': 'Doanh số',
    'DoanhSoLuyKeDong': 'Doanh số LK',
    'SoDu': 'Số dư công nợ phải thu',
}
AGING_HEADERS = {
    'TrongHan': 'Trong hạn',
    'QuaHanDuoi6T': 'Quá hạn < 6T',
    'QuaHan6Den12T': 'Quá hạn 6T-12T',
    'QuaHan12Den24T': 'Quá hạn 12T-24T',
    'QuaHan24Den36T': 'Quá hạn 24T-36T',
    'QuaHanTren36T': 'Quá hạn > 36T',
}
LABEL_COL = 4
TOLERANCE = 0.000002  # Đối chiếu trên số nguồn (triệu đồng, tấn).
DONG_TOLERANCE = 2  # Sai số làm tròn khi so số đồng đã đổi.


def _columns(source):
    """Tìm cột theo tiêu đề hàng 'Diễn giải' và hàng phân loại tuổi nợ bên dưới."""
    sh = source.sheet(SOURCE_SHEET)
    header = source.find(SOURCE_SHEET, 'Diễn giải', LABEL_COL, end=15)
    columns = {}
    for row, labels in ((header, MAIN_HEADERS), (header + 2, AGING_HEADERS)):
        for name, label in labels.items():
            hits = [c for c in range(LABEL_COL + 1, sh.max_column + 1)
                    if (key(sh.cell(row, c).value).startswith(key(label)) if name == 'DoanhSoLuyKeDong'
                        else key(sh.cell(row, c).value) == key(label))]
            if len(hits) != 1:
                raise DataError(f'Tiêu đề {label!r} cần duy nhất, thấy cột {hits}: '
                                f'{source.loc(SOURCE_SHEET, row)}')
            columns[name] = hits[0]
    source.expect(SOURCE_SHEET, header + 2, columns['SanLuongBanHang'], '(tấn)')
    if not any('triệu đồng' in key(sh.cell(r, c).value)
               for r in range(1, header) for c in range(1, sh.max_column + 1)):
        raise DataError(f'Không thấy đơn vị Triệu đồng: {source.loc(SOURCE_SHEET, f"1:{header - 1}")}')
    return header, columns


def _values(source, row, columns, strict):
    return {name: source.number(SOURCE_SHEET, row, col, blank_zero=not strict)
            for name, col in columns.items()}


DOMESTIC = ('Việt Nam', 'Nội địa')
ONSITE = ('Xuất khẩu tại chỗ', 'Xuất khẩu')
VALUE_NAMES = (*MAIN_HEADERS, *AGING_HEADERS)
CARRY_FORWARD_NAMES = ('DoanhSoLuyKeDong', 'SoDu', *AGING_HEADERS)


def _convert(values):
    # Lượng: tấn, 6 chữ số. Tiền: triệu đồng → đồng nguyên như các sheet khác.
    return {name: rounded(value, 6) if name == 'SanLuongBanHang' else dong(value)
            for name, value in values.items()}


def extract(source, report) -> dict:
    """Đọc và đối chiếu một tháng; assemble() ghi dòng khi đã biết nước của cả năm.

    Trả mọi dòng nước có trong form, kể cả dòng bằng 0, để phân biệt
    "nước có dòng nhưng không phát sinh" với "form tháng đó không có dòng nước".
    Dừng nếu tổng không khớp các phần cấu thành.
    """
    sh = source.sheet(SOURCE_SHEET)
    header, columns = _columns(source)
    total_row = source.find(SOURCE_SHEET, 'TỔNG CỘNG', LABEL_COL)
    domestic_row = source.find(SOURCE_SHEET, 'Thị trường nội địa', LABEL_COL)
    export_row = source.find(SOURCE_SHEET, 'Thị trường xuất khẩu', LABEL_COL)
    onsite_row = source.find(SOURCE_SHEET, 'XKTC', LABEL_COL)
    if not header < total_row < domestic_row < export_row < onsite_row:
        raise DataError(f'Thứ tự khối công nợ thay đổi: {source.loc(SOURCE_SHEET, "D:D")}')
    for row in (domestic_row, export_row):
        source.expect(SOURCE_SHEET, row, 3, 'SPH')
        source.validate_date(SOURCE_SHEET, f'B{row}')

    total, domestic, export = (_values(source, row, columns, True)
                               for row in (total_row, domestic_row, export_row))
    month = f'T{source.month}'
    for name, values, row in (('Tổng cộng', total, total_row), ('Nội địa', domestic, domestic_row),
                              ('Xuất khẩu', export, export_row)):
        report.check(f'{month}/Công nợ/{name}/số dư = tổng tuổi nợ',
                     sum(values[c] for c in AGING_HEADERS), values['SoDu'], TOLERANCE,
                     source.loc(SOURCE_SHEET, row, columns['SoDu']))
    for name in columns:
        report.check(f'{month}/Công nợ/{name}/tổng = nội địa + xuất khẩu',
                     domestic[name] + export[name], total[name], TOLERANCE,
                     source.loc(SOURCE_SHEET, total_row, columns[name]))

    # Nguồn chỉ ghi tuổi nợ chi tiết ở cột Trong hạn; ô quá hạn trống được coi là 0
    # và phải cộng đúng dòng nhóm, nên quá hạn phát sinh mà chưa chia nước sẽ dừng.
    blocks = {}
    for name, parent, start, end in (('Nội địa', domestic, domestic_row + 1, export_row),
                                     ('Xuất khẩu', export, export_row + 1, sh.max_row + 1)):
        rows = []
        for row in range(start, end):
            label = sh.cell(row, LABEL_COL).value
            if 'phần i -' in key(label):
                break
            if isinstance(sh.cell(row, 1).value, (int, float)) or row == onsite_row:
                if not country_name(label):
                    raise DataError(f'Dòng chi tiết thiếu diễn giải: {source.loc(SOURCE_SHEET, row, LABEL_COL)}')
                rows.append((row, _values(source, row, columns, False)))
            elif any(source.value(SOURCE_SHEET, row, c) not in (None, '', '-', 0) for c in columns.values()):
                # Dòng có số nhưng thiếu STT sẽ không được cộng: dừng thay vì bỏ sót nước.
                raise DataError(f'Dòng có số nhưng không có STT, không xác định được nước: '
                                f'{source.loc(SOURCE_SHEET, row)}')
        for column in columns:
            report.check(f'{month}/Công nợ/{name}/{column}/tổng chi tiết',
                         sum(values[column] for _, values in rows),
                         parent[column], TOLERANCE, source.loc(SOURCE_SHEET, start, columns[column]))
        blocks[name] = rows

    # Nội địa lấy dòng nhóm (gộp khách hàng nội bộ và khác) như PhanTichTieuThu.
    segments = {DOMESTIC: (domestic, source.loc(SOURCE_SHEET, domestic_row))}
    for row, values in blocks['Xuất khẩu']:
        original = text(sh.cell(row, LABEL_COL).value)
        segment = ONSITE if row == onsite_row else (country_name(original), 'Xuất khẩu')
        if segment in segments:
            raise DataError(f'Quốc gia trùng {segment[0]}: {source.loc(SOURCE_SHEET, row, LABEL_COL)}')
        if segment != ONSITE and segment[0] != original:
            report.warn('COUNTRY_NAME_ALIAS', f'{month}: đổi tên nước {original!r} thành {segment[0]!r} '
                        'để ghép với các tháng khác.', source.loc(SOURCE_SHEET, row, LABEL_COL))
        segments[segment] = (values, source.loc(SOURCE_SHEET, row))

    if total['SoDu'] and all(abs(v['SoDu'] - v['DoanhSoThangDong']) <= TOLERANCE
                             for v in (total, domestic, export)):
        report.warn('RECEIVABLE_EQUALS_MONTHLY_SALES',
                    f'{month}: số dư công nợ bằng đúng doanh số tháng ở cả nội địa và xuất khẩu; '
                    'cần xác nhận nguồn không chép nhầm cột doanh số.',
                    source.loc(SOURCE_SHEET, total_row, columns['SoDu']))
    _compare_sales_summary(source, report, _convert(total), total_row, columns)
    return dict(year=source.year, month=source.month, segments=segments,
                total=_convert(total), path=str(source.path))


def assemble(months, report):
    """Mỗi tháng xuất đủ Việt Nam, mọi nước có số khác 0 ở bất kỳ tháng nào trong năm, và XKTC.

    Nước không phát sinh trong tháng vẫn có dòng (giá trị 0) để lũy kế và
    số dư theo thời gian không bị đứt. Nếu form tháng không còn dòng của nước
    đã có lũy kế/số dư ở tháng trước thì dừng, không giả định bằng 0.
    """
    order = {}
    for month in months:
        for position, segment in enumerate(month['segments']):
            order.setdefault(segment, position)
    active = [s for s in order if s in (DOMESTIC, ONSITE)
              or any(any(m['segments'][s][0].values()) for m in months if s in m['segments'])]
    active.sort(key=lambda s: (s == ONSITE, s != DOMESTIC, order[s]))

    previous = {}
    previous_names = None
    for month in months:
        period = f"T{month['month']}"
        location_base = f"{month['path']} :: {SOURCE_SHEET}"
        names = {s[0] for s in month['segments'] if s not in (DOMESTIC, ONSITE)}
        if previous_names is not None and names != previous_names:
            report.warn('RECEIVABLE_COUNTRY_LIST_CHANGED',
                        f'{period}: danh sách nước trong form thay đổi; bỏ {sorted(previous_names - names)}, '
                        f'thêm {sorted(names - previous_names)}.', location_base)
        previous_names = names
        totals = dict.fromkeys(VALUE_NAMES, 0)
        carried_forward = dict.fromkeys(CARRY_FORWARD_NAMES, 0)
        for segment in active:
            if segment in month['segments']:
                values, location = month['segments'][segment]
                output = _convert(values)
            else:
                prior = previous.get(segment)
                output = dict.fromkeys(VALUE_NAMES, 0)
                location = f'{location_base} (không có dòng {segment[0]}; ghi 0)'
                prior_values = {
                    name: prior[name] for name in CARRY_FORWARD_NAMES
                } if prior else {}
                if any(prior_values.values()):
                    output.update(prior_values)
                    carried_forward.update({
                        name: carried_forward[name] + prior_values[name]
                        for name in CARRY_FORWARD_NAMES
                    })
                    report.warn(
                        'RECEIVABLE_COUNTRY_ROW_CARRIED_FORWARD',
                        f'{period}: thiếu dòng {segment[0]}; phát sinh tháng bằng 0, '
                        'giữ nguyên lũy kế doanh số và số dư/tuổi nợ kỳ trước. '
                        f'Tổng nguồn tháng này không có dòng chuyển tiếp: {prior_values}.',
                        location,
                    )
                else:
                    report.warn('RECEIVABLE_COUNTRY_ROW_MISSING',
                                f'{period}: form không có dòng {segment[0]}; tháng trước không có lũy kế/số dư, ghi 0.',
                                location)
            previous[segment] = output
            for name in VALUE_NAMES:
                totals[name] += output[name]
            report.add(OUTPUT_SHEET,
                       [period, *segment, output['SanLuongBanHang'], output['DoanhSoThangDong'],
                        output['DoanhSoLuyKeDong'], *(output[name] for name in AGING_HEADERS)],
                       location)
        # Nước không xuất đều bằng 0 nên các dòng xuất phải cộng ra TỔNG CỘNG
        # (tiền: sai số làm tròn tối đa 1 đồng mỗi dòng).
        for name in VALUE_NAMES:
            tolerance = TOLERANCE if name == 'SanLuongBanHang' else len(active)
            expected = month['total'][name] + carried_forward.get(name, 0)
            report.check(f'{period}/Công nợ/{name}/tổng dòng xuất = TỔNG CỘNG + chuyển tiếp',
                         totals[name], expected, tolerance, location_base)


def _compare_sales_summary(source, report, output, row, columns):
    """So với 1. KQ SXKD cùng file (đã tổng hợp trước). Lệch chỉ cảnh báo."""
    summary = {(r[1], r[2]): r[3] for r in report.rows['TongHopKQBanHang'] if r[0] == source.year}
    for period, name, labels, tolerance in (
        (f'T{source.month}', 'SanLuongBanHang', ('Số lượng bán nội địa', 'Số lượng bán xuất khẩu'), 0.001),
        (f'T{source.month}', 'DoanhSoThangDong', ('Doanh thu nội địa', 'Doanh thu xuất khẩu'), DONG_TOLERANCE),
        (f'LK{source.month}T', 'DoanhSoLuyKeDong', ('Doanh thu nội địa', 'Doanh thu xuất khẩu'), DONG_TOLERANCE),
    ):
        if any((period, label) not in summary for label in labels):
            continue
        expected = sum(summary[period, label] for label in labels)
        if abs(expected - output[name]) > tolerance:
            report.warn('RECEIVABLE_SALES_MISMATCH',
                        f'{period}/{name}: 2. Phai thu = {output[name]:,}, '
                        f'1. KQ SXKD = {expected:,} (lệch {output[name] - expected:,}).',
                        source.loc(SOURCE_SHEET, row, columns[name]))
