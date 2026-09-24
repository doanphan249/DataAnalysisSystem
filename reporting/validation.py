"""Kiểm tra khóa, số hữu hạn và quan hệ danh mục trước khi ghi file."""
from collections import Counter
import math
from .models import DataError

KEYS={
 'KyQuanTri':['Ky'], 'DonViNhaMay':['MaDonVi','NhaMay'],
 'HieuQuaSanXuat':['ChiTieu','Ky'],
 'ThongTinNhaMay':['NhaMay','Ky','ThongTinSanXuat'],
 'KetQuaSanXuat':['MaDonVi','ChiTieu','Ky'],
 'KetQuaSanXuatNhaMay':['ChiTieu','NhaMay','Ky'],
 'SanLuongSanXuat':['MatHang','NhaMay','Ky'],
 'VatTuPhuTung':['CongDoan','NhaMay','Ky'],
 'Dien':['NhaMay','NoiDung','Ky'],
 'TonKhoNguyenLieu':['Ky','LoaiTonKho','ChiTiet'],
 'TonKhoNguyenLieuChiTiet':['Ky','LoaiTonKho','NguonGoc','ChiTietTonKho'],
 'TonKhoThanhPham':['Ky','LoaiSoi'],
 'TieuThuSoi':['Ky','LoaiSoi','MatHang'],
 'PhanTichTieuThu':['Ky','QuocGia','KenhTieuThu'],
 'TongHopKQBanHang':['Nam','Ky','ChiTieu'],
 'CongNoPhaiThu':['Ky','QuocGia','KenhTieuThu'],
 'PheLieuThanhLy':['Ky','LoaiPhe'],
 'TaiChinhNhaMay':['NhaMay','Ky','ChiTieu'],
}


def validate(report, extra_periods=()):
    """Ghi chú triển khai:
- Hợp đồng không có số hợp đồng: không gộp/xóa giao dịch giống nhau.
- extra_periods: kỳ có trong dữ liệu nhưng cố ý không đưa vào KyQuanTri (T12/LK12T năm trước)."""
    periods = {r[0] for r in report.rows['KyQuanTri']} | set(extra_periods)
    plants = {r[1] for r in report.rows['DonViNhaMay']}
    for name, rows in report.rows.items():
        headers = report.schemas[name]
        if name in KEYS:
            indices = [headers.index(h) for h in KEYS[name]]
            keys = [tuple((r[i] for i in indices)) for r in rows]
            duplicates = [k for k, c in Counter(keys).items() if c > 1]
            if duplicates:
                raise DataError(f'Trùng khóa {name}: {duplicates[:5]}')
            if any((any((v in (None, '') for v in k)) for k in keys)):
                raise DataError(f'Thiếu thành phần khóa: {name}')
        for i, row in enumerate(rows, 2):
            for value in row:
                if isinstance(value, float) and (not math.isfinite(value)):
                    raise DataError(f'Số không hữu hạn {name}, hàng {i}')
            if 'Ky' in headers and row[headers.index('Ky')] not in periods:
                raise DataError(f'Kỳ không có trong danh mục: {name}, hàng {i}')
            if name != 'DonViNhaMay' and 'NhaMay' in headers and row[headers.index('NhaMay')] not in plants:
                raise DataError(f'Nhà máy không có trong danh mục: {name}, hàng {i}')
    inventory_keys = {tuple(r[:2]) for r in report.rows['TonKhoThanhPham']}
    for row in report.rows['TieuThuSoi']:
        if tuple(row[:2]) not in inventory_keys:
            raise DataError(f'Loại sợi tiêu thụ không có trong tồn kho: {row[:2]}')
    contracts = Counter((tuple(r) for r in report.rows['HDBong']))
    for identity, count in contracts.items():
        if count > 1:
            report.warn('SIMILAR_CONTRACTS', f'{count} giao dịch có cùng thông tin {identity}; giữ từng hàng nguồn.')
