"""Kế hoạch/thực hiện, thông tin sản xuất và chi phí toàn công ty/nhà máy."""
from .common import key, text, dong, rounded
from .models import DataError

ACTUAL = '1. QTSX'
PLAN = '0. Ke hoach'
METRICS = [
    ('Quy mô sản xuất','Cọc sợi'), ('Số lượng lao động','Người'),
    ('Sản lượng sản xuất','Tấn'), ('Chi số trung bình',None),
    ('Doanh thu sản xuất','Đồng'), ('Chi phí sản xuất','Đồng'),
    ('Chi phí ngoài sản xuất','Đồng'), ('Lợi nhuận trước thuế','Đồng'),
    ('Giá bông đưa vào sản xuất','USD/kg'), ('Giá xơ đưa vào sản xuất','USD/kg'),
    ('Sản lượng điện tiêu thụ','Nghìn kWh'),
]
PRICES = {m for m,_ in METRICS if m.startswith('Giá ')}
ADDITIVE = {'Sản lượng sản xuất','Doanh thu sản xuất','Chi phí sản xuất',
            'Chi phí ngoài sản xuất','Lợi nhuận trước thuế','Sản lượng điện tiêu thụ'}
PLANT_INFO = [
    ('Quy mô sản xuất thiết kế', 'Cọc sợi', 'Quy mô sản xuất thiết kế'),
    ('Quy mô sản xuất thực tế', 'Cọc sợi', 'Quy mô sản xuất thực tế'),
    ('Số lượng lao động', 'Người', 'Số lượng lao động'),
    ('Sản lượng sản xuất', 'Tấn', 'Sản lượng sản xuất'),
    ('Sản lượng điện tiêu thụ', 'Nghìn kWh', 'Sản lượng điện tiêu thụ'),
    ('Chi số trung bình', 'Ne', 'Chi số trung bình'),
    ('Giá bông đưa vào sản xuất', 'USD/Kg', 'Giá bông đưa vào sản xuất'),
    ('Giá xơ đưa vào sản xuất', 'USD/Kg', 'Giá xơ đưa vào sản xuất'),
    ('Tỷ lệ tiêu hao CD', 'Lần', 'Tỷ lệ tiêu hao bông dùng cho sợi chải thô (CD)'),
    ('Tỷ lệ tiêu hao CM', 'Lần', 'Tỷ lệ tiêu hao bông dùng cho sợi chải kỹ (CM)'),
    ('Tỷ lệ tiêu hao PE', 'Lần', 'Tỷ lệ tiêu hao xơ'),
]
# Nguyên liệu tiêu hao theo nhà máy (tấn); mẫu xếp các dòng này cuối ThongTinNhaMay.
CONSUMPTION = ['Bông tiêu hao', 'Xơ tiêu hao']
# Chỉ tiêu TaiChinhNhaMay (nguồn Trđ): tên, là số dư cuối kỳ?
PLANT_FINANCE = [
    ('Doanh thu sản xuất', False), ('Lợi nhuận trước thuế', False),
    ('Nguyên giá TSCĐ', True), ('Giá trị còn lại TSCĐ', True),
]
# Tên đích, tên nguồn, thuộc sản xuất?, biến phí?
COSTS = [
 ('Nguyên liệu Bông','Chi phí nguyên liệu chính - bông',True,True),
 ('Nguyên liệu Xơ','Chi phí nguyên liệu chính - xơ',True,True),
 ('Phế liệu thu hồi','Chi phí nguyên liệu chính - phế liệu thu hồi thanh lý (số âm)',True,True),
 ('Vật liệu phụ','Chi phí vật liệu phụ',True,True),
 ('Nhân công nhà máy','Chi phí nhân công Nhà Máy',True,True),
 ('Trích trước lương tháng 13','Chi phí trích trước lương T13',True,True),
 ('Năng lượng','Chi phí năng lượng',True,True),
 ('Vật tư phụ tùng','Chi phí vật tư phụ tùng',True,True),
 ('Khấu hao TSCĐ','Chi phí khấu hao TSCĐ',True,False),
 ('Lãi vay ngắn hạn','Chi phí Lãi vay ngắn hạn',False,True),
 ('Lãi vay dài hạn','Chi phí Lãi vay dài hạn',False,False),
 ('Hoa hồng bán hàng','Chi phí bán hàng - hoa hồng',False,True),
 ('Chi phí bán hàng khác','Chi phí bán hàng - khác',False,True),
 ('Chi phí quản lý','Chi phí quản lý',False,False),
]


def factory_columns(source, heading):
    """Ghi chú triển khai:
- Mỗi cặp là tháng + lũy kế. Nhãn NM1/NM2 ở hàng 5."""
    columns = [c.column for c in source.sheet(ACTUAL)[6] if key(c.value) == key(heading)]
    if len(columns) < 2:
        raise DataError(f'Thiếu khối nhà máy {heading}: {source.path}')
    for index, col in enumerate(columns[:2], 1):
        if key(source.sheet(ACTUAL).cell(5, col).value).replace(':', '') != f'nm{index}':
            raise DataError(f'Sai thứ tự nhà máy {heading}: {source.loc(ACTUAL, 5, col)}')
    return columns[:2]


def extract(source, report, overrides, include_previous: bool, apply_overrides: bool):
    """Ghi chú triển khai:
- Tìm tháng kế hoạch từ ngày ở dòng tiêu đề, không đoán chỉ dựa vị trí cột.
- Dòng chi phí con thường để trống đơn vị và kế thừa dòng tổng."""
    source.validate_date(ACTUAL, 'B4')
    for col, label in [(2, 'CHỈ TIÊU'), (3, 'ĐVT'), (5, 'THỰC HIỆN THÁNG TRƯỚC'), (7, 'THỰC HIỆN THÁNG BÁO CÁO')]:
        source.expect(ACTUAL, 6, col, label)
    current = factory_columns(source, 'THÁNG BÁO CÁO')
    m = source.month
    plan_columns = {}
    for cell in source.sheet(PLAN)[5]:
        value = cell.value
        if hasattr(value, 'year') and value.year == source.year:
            plan_columns[value.month] = cell.column
    if any((i not in plan_columns for i in range(1, m + 1))):
        raise DataError(f'Thiếu tiêu đề tháng kế hoạch: {source.path}')
    prod_row = source.find(PLAN, 'Sản lượng sản xuất', 3)
    weights = [source.number(PLAN, prod_row, plan_columns[i]) for i in range(1, m + 1)]
    for name, unit in METRICS:
        pr = source.find(PLAN, name, 3)
        ar = source.find(ACTUAL, 'Quy mô sản xuất thực tế' if name == 'Quy mô sản xuất' else name, 2)
        expected_unit = 'Trđ' if unit == 'Đồng' else unit
        if expected_unit:
            source.ensure_unit(PLAN, pr, 4, expected_unit)
            source.ensure_unit(ACTUAL, ar, 3, expected_unit)
        values = [source.number(PLAN, pr, plan_columns[i]) for i in range(1, m + 1)]
        if name in ADDITIVE:
            cumulative = sum(values)
        elif name == 'Chi số trung bình':
            if sum(weights) <= 0:
                raise DataError(f'Không thể gia quyền chi số: sản lượng kế hoạch <=0 tại {source.path}')
            cumulative = sum((v * w for v, w in zip(values, weights))) / sum(weights)
        elif name in PRICES:
            cumulative = sum(values) / len(values)
        else:
            cumulative = values[-1]
        plan_month = values[-1]
        for override in overrides['plan_prices'] if apply_overrides else []:
            if (source.year, m, name) == (override['year'], override['month'], override['metric']):
                if abs(plan_month - override['expected_source']) < 1e-09:
                    plan_month = override['replacement']
                    report.warn('REFERENCE_PLAN_OVERRIDE', override['reason'], source.loc(PLAN, pr, plan_columns[m]))
                else:
                    report.warn('OVERRIDE_NOT_APPLIED', 'Nguồn giá đã thay đổi; bỏ ngoại lệ giá trong mẫu.', source.loc(PLAN, pr, plan_columns[m]))
        for offset, ky, pv in [(0, f'T{m}', plan_month), (1, f'LK{m}T', cumulative)]:
            av = source.value(ACTUAL, ar, 7 + offset)
            if av is None and name in PRICES:
                candidates = [source.number(ACTUAL, ar, c + offset, blank_zero=True) for c in current]
                candidates = [v for v in candidates if v > 0]
                if not candidates:
                    raise DataError(f'Không có giá nhà máy: {source.loc(ACTUAL, ar)}')
                av = sum(candidates) / len(candidates)
                report.warn('SIMPLE_AVERAGE_PRICE', f'{name} {ky}: bình quân đơn giản giá dương nhà máy.', source.loc(ACTUAL, ar))
            else:
                av = source.number(ACTUAL, ar, 7 + offset)
            if unit == 'Đồng':
                pv, av = (dong(pv), dong(av))
            if name in PRICES:
                pv, av = (rounded(pv, 2), rounded(av, 2))
            report.add('HieuQuaSanXuat', [name, ky, pv, av], source.loc(PLAN, pr) + '; ' + source.loc(ACTUAL, ar, 7 + offset))
        if name == 'Số lượng lao động':
            monthly = source.number(ACTUAL, ar, 7)
            cum = source.number(ACTUAL, ar, 8)
            if monthly > 0 and cum > monthly * 1.5:
                report.warn('LABOR_CUMULATIVE_OUTLIER', f'Lao động T{m}={monthly}; LK{m}T={cum}. Giữ số nguồn.', source.loc(ACTUAL, ar, 8))
    for name, _, source_name in PLANT_INFO:
        row = source.find(ACTUAL, source_name, 2)
        for index, col in enumerate(current, 1):
            for offset, ky in [(0, f'T{m}'), (1, f'LK{m}T')]:
                report.add('ThongTinNhaMay', [f'SPH-{index}', ky, name, source.value(ACTUAL, row, col + offset)],
                           source.loc(ACTUAL, row, col + offset))
    _plant_consumption(source, report, current)
    _plant_finance(source, report, current)
    periods = []
    if include_previous and m == 1:
        periods.append(('T12', 'LK12T', 5, factory_columns(source, 'THÁNG TRƯỚC')))
    periods.append((f'T{m}', f'LK{m}T', 7, current))
    for month_ky, cum_ky, total_col, factory_cols in periods:
        for offset, ky in enumerate([month_ky, cum_ky]):
            group_sums = {True: 0.0, False: 0.0}
            # Mẫu xếp KetQuaSanXuatNhaMay theo nhà máy rồi chỉ tiêu trong mỗi kỳ.
            plant_rows = [[] for _ in factory_cols]
            for name, label, is_production, is_variable in COSTS:
                if name == 'Trích trước lương tháng 13':
                    candidates = [text(r[1].value) for r in source.sheet(ACTUAL) if key(r[1].value).startswith(key(label))]
                    if len(candidates) != 1:
                        raise DataError(f'Không nhận diện lương T13: {source.path}')
                    label = candidates[0]
                row = source.find(ACTUAL, label, 2)
                unit_row = row if text(source.sheet(ACTUAL).cell(row, 3).value) else source.find(ACTUAL, 'Chi phí sản xuất' if is_production else 'Chi phí ngoài sản xuất', 2)
                source.ensure_unit(ACTUAL, unit_row, 3, 'Trđ')
                raw = source.number(ACTUAL, row, total_col + offset, blank_zero=True)
                group_sums[is_production] += raw
                group = 'Chi phí sản xuất' if is_production else 'Chi phí ngoài sản xuất'
                report.add('KetQuaSanXuat', ['SPH', name, group, 'Biến phí' if is_variable else 'Định phí', ky, dong(raw)], source.loc(ACTUAL, row, total_col + offset))
                plant_values = []
                for index, col in enumerate(factory_cols, 1):
                    value = source.number(ACTUAL, row, col + offset, blank_zero=True)
                    plant_values.append(value)
                    target_name = 'Chi phí quản lí' if name == 'Chi phí quản lý' else name
                    plant_rows[index - 1].append(([target_name, f'SPH-{index}', ky, dong(value)], source.loc(ACTUAL, row, col + offset)))
                report.check(f'{ky}/{name}/hai nhà máy', sum(plant_values), raw, 1e-06, source.loc(ACTUAL, row))
            for rows in plant_rows:
                for values, location in rows:
                    report.add('KetQuaSanXuatNhaMay', values, location)
            for is_production, label in [(True, 'Chi phí sản xuất'), (False, 'Chi phí ngoài sản xuất')]:
                row = source.find(ACTUAL, label, 2)
                report.check(f'{ky}/{label}/chi tiết', group_sums[is_production], source.number(ACTUAL, row, total_col + offset), 1e-06, source.loc(ACTUAL, row))


def _plant_consumption(source, report, current):
    """Bông/Xơ tiêu hao theo nhà máy (tấn), giữ số nguồn như các dòng ThongTinNhaMay khác.

    Ghi chú triển khai:
- Thứ tự trong tháng theo mẫu: kỳ → chỉ tiêu → nhà máy."""
    m = source.month
    rows = {name: source.find(ACTUAL, name, 2) for name in CONSUMPTION}
    for row in rows.values():
        source.ensure_unit(ACTUAL, row, 3, 'Tấn')
    for offset, ky in [(0, f'T{m}'), (1, f'LK{m}T')]:
        for name, row in rows.items():
            values = [source.number(ACTUAL, row, col + offset, blank_zero=True) for col in current]
            for index, (col, value) in enumerate(zip(current, values), 1):
                report.add('ThongTinNhaMay', [f'SPH-{index}', ky, name, value], source.loc(ACTUAL, row, col + offset))
            report.check(f'{ky}/{name}/hai nhà máy', sum(values), source.number(ACTUAL, row, 7 + offset, blank_zero=True), 1e-06, source.loc(ACTUAL, row))


def _plant_finance(source, report, current):
    """Doanh thu, lợi nhuận và TSCĐ theo nhà máy cho TaiChinhNhaMay; Trđ → đồng.

    Ghi chú triển khai:
- Nguồn chỉ ghi số dư TSCĐ ở cột lũy kế (cuối kỳ); dùng số đó cho cả Tn và LKnT.
- Thứ tự trong tháng theo mẫu: kỳ → nhà máy → chỉ tiêu."""
    m = source.month
    rows = {name: source.find(ACTUAL, name, 2) for name, _ in PLANT_FINANCE}
    for row in rows.values():
        source.ensure_unit(ACTUAL, row, 3, 'Trđ')
    for offset, ky in [(0, f'T{m}'), (1, f'LK{m}T')]:
        values = {}
        for name, is_balance in PLANT_FINANCE:
            row, shift = rows[name], 1 if is_balance else offset
            values[name] = [(source.number(ACTUAL, row, col + shift), source.loc(ACTUAL, row, col + shift)) for col in current]
            report.check(f'{ky}/{name}/hai nhà máy', sum(v for v, _ in values[name]), source.number(ACTUAL, row, 7 + shift), 1e-06, source.loc(ACTUAL, row))
        for index in range(len(current)):
            for name, _ in PLANT_FINANCE:
                value, location = values[name][index]
                report.add('TaiChinhNhaMay', [f'SPH-{index + 1}', ky, name, dong(value)], location)
