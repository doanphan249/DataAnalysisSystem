"""Điều phối toàn bộ quá trình. Không đọc các workbook kết quả cũ."""
from pathlib import Path
from datetime import datetime
import hashlib
import json
import traceback
from .models import Report,RunResult,DataError
from .sources import Source,discover
from . import production,operations,inventory,sales,sales_summary,receivables
from .validation import validate
from .excel_writer import write_report

CONFIG=Path(__file__).parent/'config'
# Kỳ năm trước lấy từ báo cáo T1: có trong dữ liệu chi phí/VTPT nhưng không vào KyQuanTri,
# vì Power BI xếp kỳ theo số tháng nên T12 năm trước sẽ đứng sau T8 năm nay.
PREVIOUS_PERIODS=('T12','LK12T')


def load_config(name):
    return json.loads((CONFIG / name).read_text(encoding='utf-8'))


def _catalogs(report, months):
    for item in months:
        report.add('KyQuanTri', [f'T{item.month}', f'Tháng {item.month}'], 'Danh mục kỳ')
    for item in months:
        report.add('KyQuanTri', [f'LK{item.month}T', f'Lũy kế {item.month} tháng'], 'Danh mục kỳ')
    for plant in [1, 2]:
        report.add('DonViNhaMay', ['SPH', f'SPH-{plant}'], 'Danh mục mẫu SPH')


def _order_rows(report):
    """Thứ tự dòng giống mẫu Dashboard khi dữ liệu trải nhiều tháng. Power BI không phụ thuộc thứ tự này."""
    report.reorder('HieuQuaSanXuat', lambda r: not r[1].startswith('LK'))
    report.reorder('VatTuPhuTung', lambda r: operations.STAGES.index(r[0]))
    report.reorder('ThongTinNhaMay', lambda r: r[2] in production.CONSUMPTION)


def run_report(settings) -> RunResult:
    """Ghi chú triển khai:
- Kiểm tra lại hash: tránh đọc một bộ nguồn đang bị chỉnh sửa trong lúc chạy.
- Thứ tự lineage tương ứng hàng Excel 2,3,... của từng sheet."""
    source_folder = settings.source_folder.resolve()
    output = settings.output_file.resolve()
    if output.suffix.lower() != '.xlsx':
        raise DataError('OUTPUT_FILE phải kết thúc bằng .xlsx')
    if output.is_relative_to(source_folder):
        raise DataError('File đích phải nằm ngoài thư mục nguồn')
    run_id = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    output.parent.mkdir(parents=True, exist_ok=True)
    log_dir = output.parent / 'logs'
    log_dir.mkdir(exist_ok=True)
    audit_file = log_dir / f'{output.stem}_{run_id}.json'
    text_log = audit_file.with_suffix('.log')
    report = Report(load_config('schemas.json'))
    audit = {'run_id': run_id, 'status': 'running', 'source_folder': str(source_folder), 'output_file': str(output), 'sources': [], 'settings': {'year': settings.year, 'include_previous_december': settings.include_previous_december, 'apply_reference_overrides': settings.apply_reference_overrides}}
    try:
        months = discover(source_folder, settings.year)
        include_previous = settings.include_previous_december and months[-1].month < 12
        if settings.include_previous_december and (not include_previous):
            report.warn('PREVIOUS_DECEMBER_OMITTED', 'Đã có T12 năm hiện tại: bỏ T12/LK12T năm trước để tránh trùng mã kỳ. Xuất riêng năm trước nếu cần.')
        audit['report_year'] = months[0].year
        audit['months'] = [m.month for m in months]
        audit['previous_december_included'] = include_previous
        overrides = load_config('reference_overrides.json')
        names = load_config('product_names.json')
        audit['config_sha256'] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in CONFIG.glob('*.json')}
        country_months = []
        receivable_months = []
        for month in months:
            print(f'Đang tổng hợp tháng {month.month:02d}/{month.year}...')
            for path in [month.production, month.sales]:
                audit['sources'].append({'path': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
            with Source(month.production, month.year, month.month, report) as source:
                source.validate_company('0. Ke hoach')
                production.extract(source, report, overrides, include_previous, settings.apply_reference_overrides)
                operations.extract(source, report, names, include_previous)
                inventory.extract(source, report, overrides)
                sales.extract_production_book(source, report, overrides, settings.apply_reference_overrides)
            with Source(month.sales, month.year, month.month, report) as source:
                source.validate_company('2. Phai thu')
                country_months.append(sales.read_country_sales(source, report))
                sales_summary.extract(source, report)
                receivable_months.append(receivables.extract(source, report))
        sales.assemble_country_sales(country_months, report)
        receivables.assemble(receivable_months, report)
        _catalogs(report, months)
        _order_rows(report)
        validate(report, PREVIOUS_PERIODS if include_previous else ())
        for entry in audit['sources']:
            if hashlib.sha256(Path(entry['path']).read_bytes()).hexdigest() != entry['sha256']:
                raise DataError(f"Nguồn thay đổi trong lúc chạy: {entry['path']}. Chạy lại khi đã lưu xong nguồn.")
        backup = write_report(report, output, run_id)
        audit['backup'] = str(backup) if backup else None
        audit['status'] = 'completed'
        audit['output_sha256'] = hashlib.sha256(output.read_bytes()).hexdigest()
    except Exception as exc:
        audit['status'] = 'failed'
        audit['error'] = str(exc)
        audit['traceback'] = traceback.format_exc()
        raise
    finally:
        audit['finished_at'] = datetime.now().isoformat()
        audit['row_counts'] = {name: len(rows) for name, rows in report.rows.items()}
        audit['warnings'] = report.warnings
        audit['checks'] = report.checks
        audit['lineage_start_excel_row'] = 2
        audit['lineage'] = report.lineage
        audit_file.write_text(json.dumps(audit, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        lines = [f"Trạng thái: {audit['status']}", f'File đích: {output}', f"Năm: {audit.get('report_year', 'chưa xác định')}", f'Số kiểm tra đối chiếu: {len(report.checks)}', f'Số cảnh báo: {len(report.warnings)}', *[f'{name}: {len(rows)} dòng' for name, rows in report.rows.items()], *[f"[{w['code']}] {w['message']} | {w['source']}" for w in report.warnings]]
        if 'error' in audit:
            lines.append('LỖI: ' + audit['error'])
        text_log.write_text('\n'.join(lines), encoding='utf-8')
    business = sum((len(rows) for name, rows in report.rows.items() if name not in ('KyQuanTri', 'DonViNhaMay')))
    return RunResult(output, audit_file, business, len(report.warnings))
