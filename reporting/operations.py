"""Chi tiết sản lượng, vật tư phụ tùng và điện."""
from .common import text, key, rounded, dong, normalize_product, classify_product
from .models import DataError

STAGES=['Điều không','Dây bông','Chải thô','Ghép','Cuộn cúi chải kỹ',
        'Thô','Sợi con','Ống','Cải tạo sửa chữa','Khác']


def extract(source, report, product_names, include_previous):
    _production(source, report, product_names)
    _spare_parts(source, report, include_previous)
    _electricity(source, report)


def _production(source, report, mapping):
    """Ghi chú triển khai:
- Hàng giữ chỗ có STT, không có sản phẩm/phát sinh."""
    sheet = '1a. SLSX'
    sh = source.sheet(sheet)
    source.validate_date(sheet, 'B4')
    header = source.find(sheet, 'Mặt hàng', 2)
    source.expect(sheet, header, 3, 'NM1')
    source.expect(sheet, header, 4, 'NM2')
    if not any(('tấn' in key(c.value) for row in sh.iter_rows(max_row=header) for c in row)):
        raise DataError(f'Không xác định được đơn vị tấn: {source.path} :: {sheet}')
    sums = [0.0, 0.0]
    count = 0
    for row in range(header + 1, sh.max_row + 1):
        seq = sh.cell(row, 1).value
        if not isinstance(seq, (int, float)) or not float(seq).is_integer() or seq < 1:
            continue
        original = text(sh.cell(row, 2).value)
        if not original:
            if any((source.number(sheet, row, c, blank_zero=True) for c in [3, 4])):
                raise DataError(f'Có sản lượng nhưng thiếu mặt hàng: {source.loc(sheet, row, 2)}')
            continue
        name = normalize_product(original, mapping)
        classification = classify_product(name)
        if not all(classification):
            report.warn('UNKNOWN_PRODUCT_CLASS', f'Chưa phân loại đủ mặt hàng: {name}', source.loc(sheet, row, 2))
        for index, col in enumerate([3, 4]):
            qty = source.number(sheet, row, col, blank_zero=True)
            sums[index] += qty
            report.add('SanLuongSanXuat', [name, *classification, f'SPH-{index + 1}', f'T{source.month}', qty], source.loc(sheet, row, col))
        count += 1
    if not count:
        raise DataError(f'Không có mặt hàng sản xuất: {source.path}')
    for index in [0, 1]:
        label = f'NM{index + 1}'
        cells = [c for row in sh.iter_rows(max_row=header - 1) for c in row if key(c.value) == key(label)]
        if len(cells) != 1:
            raise DataError(f'Thiếu tổng {label}: {source.path} :: {sheet}')
        c = cells[0]
        report.check(f'T{source.month}/SLSX/{label}', sums[index], source.number(sheet, c.row, c.column + 1), 1e-06, source.loc(sheet, c.row, c.column + 1))


def _spare_parts(source, report, include_previous):
    """Ghi chú triển khai:
- Cột B/C kỳ trước, E/F kỳ báo cáo; xác nhận tên nhà máy và tháng."""
    sheet = '1b. VTPT'
    source.validate_date(sheet, 'A4')
    if 'triệu đồng' not in key(source.sheet(sheet)['G4'].value):
        raise DataError(f'Đơn vị VTPT không phải triệu đồng: {source.path}')
    start = source.find(sheet, STAGES[0], 1)
    for col in [2, 5]:
        source.expect(sheet, start - 1, col, 'NM1')
    for col in [3, 6]:
        source.expect(sheet, start - 1, col, 'NM2')
    source.expect(sheet, start - 2, 2, 'THÁNG TRƯỚC')
    source.expect(sheet, start - 2, 5, 'THÁNG BÁO CÁO')
    periods = [(f'T{source.month}', 5)]
    if include_previous and source.month == 1:
        periods.insert(0, ('T12', 2))
    for ky, col in periods:
        totals = [0, 0]
        for offset, stage in enumerate(STAGES):
            row = start + offset
            source.expect(sheet, row, 1, stage)
            for plant in [0, 1]:
                value = dong(source.number(sheet, row, col + plant, blank_zero=True))
                totals[plant] += value
                report.add('VatTuPhuTung', [stage, f'SPH-{plant + 1}', ky, value], source.loc(sheet, row, col + plant))
        total_row = start + len(STAGES)
        if 'tổng' not in key(source.sheet(sheet).cell(total_row, 1).value):
            raise DataError(f'Thiếu dòng tổng VTPT: {source.loc(sheet, total_row)}')
        for plant in [0, 1]:
            report.check(f'{ky}/VTPT/SPH-{plant + 1}', totals[plant], dong(source.number(sheet, total_row, col + plant)), 2, source.loc(sheet, total_row, col + plant))


def _electricity(source, report):
    """Ghi chú triển khai:
- Nhãn khối ở C, nguồn T8 dùng NM1/NM2 thay vì SPH-1/SPH-2."""
    sheet = '1d. Dien'
    sh = source.sheet(sheet)
    source.validate_date(sheet, 'A5')
    for col, unit in [(5, 'nghìn kwh'), (6, 'đồng/ kwh'), (7, 'trđ')]:
        if unit not in key(sh.cell(7, col).value):
            raise DataError(f'Đơn vị điện thay đổi: {source.loc(sheet, 7, col)}')
    starts = [r for r in range(2, sh.max_row + 1) if key(sh.cell(r, 4).value) == key('Điện lưới') and key(sh.cell(r - 1, 4).value).startswith('nm')]
    if len(starts) < 2:
        raise DataError(f'Cần 2 khối điện nhà máy: {source.path}')
    for extra in starts[2:]:
        if any((source.number(sheet, extra - 1, c, blank_zero=True) for c in [5, 7])):
            raise DataError(f'Có nhà máy ngoài SPH-1/SPH-2 phát sinh điện: {source.loc(sheet, extra - 1)}')
    starts = starts[:2]
    contents = ['Điện lưới', 'Điện tái tạo', 'Chi phí khấu hao']
    for plant, start in enumerate(starts, 1):
        totals = [0.0, 0.0]
        block = text(sh.cell(start - 1, 3).value)
        if block in ('…', '...', '') and key(sh.cell(start - 1, 4).value) == f'nm{plant}':
            report.warn('FACTORY_PLACEHOLDER', f'Mã nhà máy {block!r}: dùng nhãn NM{plant} để ánh xạ SPH-{plant}.', source.loc(sheet, start - 1, 3))
            block = f'NM{plant}'
        if key(block) not in {key(f'SPH-{plant}'), key(f'NM{plant}')}:
            raise DataError(f'Không nhận diện nhà máy điện: {source.loc(sheet, start - 1, 3)} = {block!r}')
        for offset, content in enumerate(contents):
            row = start + offset
            source.expect(sheet, row, 4, content)
            qty = source.number(sheet, row, 5, blank_zero=True)
            amount = source.number(sheet, row, 7, blank_zero=True)
            report.add('Dien', [f'SPH-{plant}', content, f'T{source.month}', rounded(qty * 1000, 4), dong(amount)], source.loc(sheet, f'E{row}:G{row}'))
            if offset < 2:
                totals[0] += qty
                totals[1] += amount
        for i, col in enumerate([5, 7]):
            report.check(f'T{source.month}/Điện/SPH-{plant}/cột {col}', totals[i], source.number(sheet, start - 1, col, blank_zero=True), 1e-07, source.loc(sheet, start - 1, col))
