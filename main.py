"""Điểm chạy duy nhất: sửa cấu hình bên dưới, sau đó chạy python src/main.py."""
from pathlib import Path
import sys

# ================= CẤU HÌNH BẠN CÓ THỂ THAY ĐỔI =================
SOURCE_FOLDER = Path(r"D:\DOAN\MonthlyReportingData")
OUTPUT_FILE = Path(r"D:\DOAN\DataBase\DuLieuBaoCao_TongHop1.xlsx")

# None: tự nhận diện nếu nguồn chỉ có một năm. Có nhiều năm: ghi năm cần lấy.
REPORT_YEAR: int | None = None

# Lấy thêm T12/LK12T năm trước từ báo cáo T1. Tự bỏ khi đã có T12 năm hiện tại
# để không trùng mã kỳ trong cấu trúc Excel không có cột năm.
INCLUDE_PREVIOUS_DECEMBER = True

# True: tái tạo các điều chỉnh của mẫu đã học, ghi rõ trong nhật ký.
# False: giữ số nguồn, không bổ sung ngày hợp đồng không có trong nguồn.
APPLY_REFERENCE_OVERRIDES = True
# ==============================================================


def main() -> int:
    try:
        from reporting.pipeline import run_report
        from reporting.models import Settings

        result = run_report(Settings(
            source_folder=SOURCE_FOLDER,
            output_file=OUTPUT_FILE,
            year=REPORT_YEAR,
            include_previous_december=INCLUDE_PREVIOUS_DECEMBER,
            apply_reference_overrides=APPLY_REFERENCE_OVERRIDES,
        ))
        print(f"\nHOÀN THÀNH: {result.output_file}")
        print(f"Số dòng nghiệp vụ: {result.business_rows:,}")
        print(f"Số cảnh báo: {result.warnings}. Xem: {result.audit_file}")
        return 0
    except ImportError as exc:
        print(f"Thiếu thư viện: {exc}\nChạy: python -m pip install -r \"{Path(__file__).with_name('requirements.txt')}\"", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"\nKHÔNG HOÀN THÀNH: {exc}", file=sys.stderr)
        print("File đích cũ được giữ nếu lỗi xảy ra trước khi xuất thành công. Xem thư mục logs cạnh file đích.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
