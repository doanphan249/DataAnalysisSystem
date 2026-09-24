# Tổng hợp dữ liệu báo cáo quản trị SPH bằng Python

Chương trình đọc các báo cáo tháng trong `MonthlyReportingData`, tổng hợp **19 sheet dữ liệu và sheet GhiChu** theo đúng mẫu `D:\Dashboard\DuLieuBaoCao.xlsx` — file mà Power BI (`BaoCaoQuanTri_SPH_PhanVanDoan`) đọc — và xuất một workbook hoàn chỉnh. Bạn chỉ chạy **`src/main.py`**. Các tác vụ còn lại được import tự động.

Hai đường dẫn cần quan tâm nằm ngay đầu `src/main.py`:

```python
SOURCE_FOLDER = Path(r"G:\DOAN\MonthlyReportingData")
OUTPUT_FILE = Path(r"G:\DOAN\DataBase\DuLieuBaoCao_TongHop1.xlsx")
```

`SOURCE_FOLDER` là **thư mục nguồn**. `OUTPUT_FILE` là **đường dẫn đầy đủ đến file Excel đích**, bao gồm tên file và đuôi `.xlsx`. Chương trình tự tạo thư mục đích nếu chưa có.

Không cần mở Excel, không cần chạy Node.js, không cần tài khoản API hay cài Power BI để chạy bộ mã này. Không cần file Excel mẫu khi chạy: cấu trúc cột và quy tắc đã được tách vào mã Python cùng các file cấu hình JSON đi kèm.

## 1. Cài đặt lần đầu trên Windows

### Bước 1 — Cài Python

Sử dụng **Python 3.10 trở lên**, khuyến nghị Python 3.12. Có thể lấy bộ cài tại [trang tải Python chính thức](https://www.python.org/downloads/windows/).

Khi cài đặt, chọn **Add python.exe to PATH**. Sau khi cài, mở cửa sổ PowerShell hoặc Terminal mới, kiểm tra:

```powershell
py --version
```

Nếu máy không có lệnh `py` nhưng có `python`, dùng `python --version` và thay `py` bằng `python` trong bước tạo môi trường bên dưới.

### Bước 2 — Mở đúng thư mục mã nguồn

Trong PowerShell:

```powershell
Set-Location "G:\DOAN"
```

Trong thư mục này cần có thư mục `src`, bên trong chứa `main.py`, `reporting`, `requirements.txt`, `README.md` và `tests`. Khi chuyển máy, sao chép toàn bộ `src`. Bộ kiểm thử nằm trong `src/tests`.

### Bước 3 — Tạo môi trường thư viện riêng

```powershell
py -m venv .venv
```

Lệnh này tạo thư mục `.venv` ngay trong `G:\DOAN`. Chỉ cần làm một lần. Không chép `.venv` từ máy khác sang; tạo lại trên máy mới.

### Bước 4 — Cài các thư viện

```powershell
.\.venv\Scripts\python.exe -m pip install -r src/requirements.txt
```

Máy cần kết nối Internet tới PyPI khi cài lần đầu. Sau khi cài xong, chạy báo cáo không cần Internet.

`src/requirements.txt` khóa các phiên bản đã dùng kiểm tra:

| Thư viện | Phiên bản | Mục đích |
|---|---|---|
| openpyxl | 3.1.5 | Đọc dữ liệu, công thức nguồn và kiểm tra workbook đã xuất |
| XlsxWriter | 3.2.9 | Tạo workbook, định dạng, bảng Excel, công thức và kết quả đã tính |

`pip` sẽ tự cài thư viện phụ thuộc như `et-xmlfile`; không cần cài từng gói bằng tay. Không cần `pandas`, `numpy`, `pytest`, `pywin32` hoặc Microsoft Office để chạy chương trình.

Hướng dẫn dùng đường dẫn trực tiếp đến Python trong `.venv`, vì vậy **không cần kích hoạt môi trường** và không cần đổi Execution Policy của PowerShell.

## 2. Chuẩn bị thư mục nguồn

Ví dụ:

```text
G:\DOAN\MonthlyReportingData\
├── BCQT_012026\
│   ├── Form BC QTSX T01.26-cap nhat gui TD.xlsx
│   └── Form BC KQ Tieu thu T01.26.xlsx
├── BCQT_022026\
│   ├── Form BC QTSX T02.26.xlsx
│   └── Form BC KQ Tieu thu T02.26.xlsx
...
└── BCQT_082026\
    ├── Phú Hưng- Form BC QTSX T08.26.xlsx
    └── Phú Hưng- Form BC KQ Tieu thu T08.26.xlsx
```

Quy tắc nhận diện:

- Tên thư mục tháng: `BCQT_MMYYYY`, ví dụ `BCQT_092026` là tháng 9/2026.
- Trong mỗi thư mục có đúng **một file `.xlsx` có `QTSX` trong tên** và **một file `.xlsx` có `KQ Tieu thu` trong tên**. Có thể có tiền tố hoặc hậu tố khác.
- Chương trình bỏ qua file khóa Excel bắt đầu bằng `~$`.
- Nếu có hai bản QTSX cho cùng tháng, chuyển bản không dùng ra ngoài thư mục tháng. Chương trình dừng để tránh tự chọn nhầm.
- Phải có đủ tháng từ tháng 1 đến tháng cuối cùng muốn tổng hợp. Ví dụ có T1 và T3 nhưng thiếu T2 thì chương trình dừng, vì không đủ dữ liệu tính sản lượng tiêu thụ lũy kế.
- File nguồn cần được tính toán và lưu bởi Excel/ứng dụng tương thích. `openpyxl` không tính lại công thức nguồn. Nếu công thức không có giá trị đã lưu, chương trình báo chính xác file/sheet/ô cần kiểm tra.
- Bộ mã dành cho đơn vị **SPH — Phú Hưng**, hai nhà máy **SPH-1 và SPH-2**, và bố cục báo cáo đã khảo sát. Không dùng nguyên trạng để xử lý đơn vị khác.

## 3. Sửa cấu hình trong main.py

Mở `src/main.py` bằng VS Code, Notepad hoặc trình soạn thảo khác. Chỉnh phần đầu:

```python
SOURCE_FOLDER = Path(r"G:\DOAN\MonthlyReportingData")
OUTPUT_FILE = Path(r"G:\DOAN\DataBase\DuLieuBaoCao_TongHop1.xlsx")
REPORT_YEAR: int | None = None
INCLUDE_PREVIOUS_DECEMBER = True
APPLY_REFERENCE_OVERRIDES = True
```

Giữ chữ `r` trước chuỗi đường dẫn Windows để dấu `\` được hiểu đúng.

| Cấu hình | Ý nghĩa |
|---|---|
| `SOURCE_FOLDER` | Nơi chứa các thư mục tháng |
| `OUTPUT_FILE` | File Excel cần tạo; phải nằm ngoài thư mục nguồn |
| `REPORT_YEAR = None` | Tự xác định năm nếu nguồn chỉ có một năm |
| `REPORT_YEAR = 2026` | Chỉ lấy các thư mục năm 2026 nếu thư mục nguồn có nhiều năm |
| `INCLUDE_PREVIOUS_DECEMBER = True` | Lấy thêm T12/LK12T năm trước từ phần kỳ trước trong báo cáo T1 cho các bảng có dữ liệu này (`KetQuaSanXuat`, `KetQuaSanXuatNhaMay`, `VatTuPhuTung`). Hai kỳ này **không** đưa vào `KyQuanTri`, vì Power BI xếp kỳ theo số tháng nên T12 năm trước sẽ đứng sau T8 năm nay |
| `APPLY_REFERENCE_OVERRIDES = True` | Áp dụng các điều chỉnh riêng của mẫu đã học: giá bông T8 và ngày hợp đồng T6, kèm nhật ký |

Mặc định xuất sang `DuLieuBaoCao_TongHop1.xlsx` vì file `DataBase\DuLieuBaoCao.xlsx` đã tồn tại và bị khóa trong lần chạy kiểm tra. Bạn có thể đổi thành tên bất kỳ, kể cả `DuLieuBaoCao.xlsx`, sau khi đóng ứng dụng đang giữ file đó.

**Một workbook chứa một năm báo cáo.** Do phần lớn sheet không có cột `Nam`, không thể trộn nhiều năm chỉ với mã kỳ `T1`, `T2`… Nếu đã có T12 của năm hiện tại, chương trình tự bỏ T12/LK12T năm trước và ghi cảnh báo, tránh hai dòng khác năm dùng cùng khóa. Dữ liệu năm trước vẫn nằm trong file nguồn; muốn lưu lịch sử hãy xuất mỗi năm thành file riêng.

## 4. Chạy chương trình

Từ `G:\DOAN`, chạy duy nhất:

```powershell
.\.venv\Scripts\python.exe .\src\main.py
```

Không cần chạy lần lượt các file trong `reporting`. `main.py` gọi `pipeline.py`, và pipeline tự điều phối các mô-đun nghiệp vụ.

Cũng có thể chạy từ thư mục bất kỳ bằng đường dẫn đầy đủ:

```powershell
& "G:\DOAN\.venv\Scripts\python.exe" "G:\DOAN\src\main.py"
```

Nếu chạy bằng nút Run trong VS Code, chọn Python interpreter `G:\DOAN\.venv\Scripts\python.exe`, rồi mở và chạy `src/main.py`.

Kết quả thành công có dạng:

```text
Đang tổng hợp tháng 01/2026...
...
HOÀN THÀNH: G:\DOAN\DataBase\DuLieuBaoCao_TongHop1.xlsx
Số dòng nghiệp vụ: 2,804
Số cảnh báo: ... Xem: ...\logs\DuLieuBaoCao_TongHop_....json
```

Mã thoát là `0` khi thành công, `1` khi có lỗi. Trong PowerShell có thể xem `$LASTEXITCODE` ngay sau lệnh chạy.

## 5. Cách sử dụng hàng tháng

1. Thêm thư mục mới, ví dụ `BCQT_092026`.
2. Đặt hai file QTSX và KQ tiêu thụ đã hoàn thành, tính lại và lưu vào thư mục đó.
3. Đóng workbook đích nếu đang mở trong Excel hoặc bị ứng dụng khác giữ khóa.
4. Chạy lại `src/main.py`.
5. Đọc dòng `HOÀN THÀNH`, kiểm tra file đích và các cảnh báo trong `logs`.

Chương trình **đọc lại tất cả các tháng đã chọn và xây lại workbook**, không cộng nối vào file kết quả cũ. Vì vậy, chạy lại không làm nhân đôi dữ liệu. Không cần đổi danh sách tháng trong mã khi thêm T9–T12 nếu nguồn giữ nguyên bố cục.

Nếu tên nguồn gốc, sản phẩm hoặc quốc gia mới xuất hiện, chương trình đọc các dòng có phát sinh. Bốn trường phân loại sản phẩm sử dụng quy tắc của mẫu; sản phẩm chưa phân loại đầy đủ sẽ được ghi cảnh báo để bạn bổ sung ánh xạ. Nếu bố cục hàng/cột hoặc đơn vị nguồn thay đổi, cần cập nhật mô-đun tương ứng sau khi kiểm tra cảnh báo/lỗi; không coi chương trình là bộ đọc mọi mẫu Excel.

## 6. Nội dung workbook đầu ra

Giữ đúng tên sheet, thứ tự sheet, tên bảng Excel, tên và thứ tự cột của mẫu `D:\Dashboard\DuLieuBaoCao.xlsx`. **Power BI đọc từng bảng theo tên bảng Excel** (ví dụ `Item="PhanTichTieuThu",Kind="Table"`), nên mọi sheet dữ liệu đều là một bảng Excel, kể cả khi không có dòng phát sinh. Hiện tại có:

| Sheet | Bảng Excel | Cột | Dòng T1–T8/2026 |
|---|---|---|---:|
| KyQuanTri | KyQuanTri | Ky, KyQuanTri | 16 |
| DonViNhaMay | DonViNhaMay | MaDonVi, NhaMay | 2 |
| HieuQuaSanXuat | HieuQuaSanXuat | ChiTieu, Ky, KeHoach, ThucHien | 176 |
| ThongTinNhaMay | ThongTinNhaMay | NhaMay, Ky, ThongTinSanXuat, GiaTri | 416 |
| KetQuaSanXuat | KetQuaSanXuat | MaDonVi, ChiTieu, NhomChiPhi, NhomPhi, Ky, GiaTri | 252 |
| KetQuaSanXuatNhaMay | KetQuaSanXuatNhaMay | ChiTieu, NhaMay, Ky, GiaTri | 504 |
| SanLuongSanXuat | SanLuongSanXuat | MatHang, TheoNguyenLieu, TheoLoaiSoi, LoaiChai, ChiSoSoi, NhaMay, Ky, SanLuong | 268 |
| VatTuPhuTung | ChiPhiVTPTTable | CongDoan, NhaMay, Ky, GiaTri | 180 |
| Dien | NangLuongDienTable | NhaMay, NoiDung, Ky, SanLuongDienTieuThu, ThanhTien | 48 |
| TonKhoNguyenLieu | TonKhoNguyenLieuTable | Ky, LoaiTonKho, ChiTiet, SoLuongTan, GiaTri | 80 |
| TonKhoNguyenLieuChiTiet | ChiTietTonKhoNLTable | Ky, LoaiTonKho, NguonGoc, ChiTietTonKho, SoLuongTan, DonGiaUSDkg, ThanhTienDong | 204 |
| TonKhoThanhPham | TonKhoThanhPhamTable | Ky, LoaiSoi, XuatBanTrongThang, TonCuoiThang, GiaTriTonDong | 64 |
| TieuThuSoi | TieuThuSoiTable | Ky, LoaiSoi, MatHang, GiaCIFUSDkg, IPI, UPercent | 41 |
| PhanTichTieuThu | PhanTichTieuThu | Ky, QuocGia, KenhTieuThu, DoanhSoDong | 144 |
| TongHopKQBanHang | TongHopKQBanHangTable | Nam, Ky, ChiTieu, GiaTri | 144 |
| PheLieuThanhLy | PheLieuThanhLyTable | Ky, LoaiPhe, SoLuong, GiaTri | 64 |
| HDBong | MuaBongTable | Ky, NgayHopDongPLHD, SoLuongTan, DonGiaUSDkg, ThanhTienUSD | 19 |
| CongNoPhaiThu | CongNoPhaiThuTable | Ky, QuocGia, KenhTieuThu, SanLuongBanHang, DoanhSoThangDong, DoanhSoLuyKeDong, TrongHan, QuaHanDuoi6T, QuaHan6Den12T, QuaHan12Den24T, QuaHan24Den36T, QuaHanTren36T | 72 |
| TaiChinhNhaMay | TaiChinhNhaMay | NhaMay, Ky, ChiTieu, GiaTri | 128 |
| GhiChu | — | Danh sách mặt hàng khác nhau và quy ước mã mặt hàng | — |

Tổng **2.804 dòng nghiệp vụ**, cộng **18 dòng danh mục**, là **2.822 dòng** không tính tiêu đề. Số dòng nghiệp vụ là đối chứng của bộ dữ liệu hiện tại, không được ép cố định cho các tháng sau.

Tên bảng khác tên sheet nằm trong `TABLE_NAMES` (`reporting/excel_writer.py`); tên cột và thứ tự cột nằm trong `config/schemas.json`. Đổi một trong hai sẽ làm Power BI không làm mới được, trừ khi sửa cả mô hình Power BI.

File được tạo mới với định dạng bảng, cố định tiêu đề và kiểu dữ liệu số/ngày. Định dạng số theo mẫu: tiền `#,##0`; cột một giá trị nhiều đơn vị (`HieuQuaSanXuat`, `ThongTinNhaMay`, `TongHopKQBanHang`) chọn định dạng theo đơn vị của từng chỉ tiêu. Công thức dùng tham chiếu có cấu trúc: phân loại ở `SanLuongSanXuat!B:E` (`SanLuongSanXuat[[#This Row],[MatHang]]`) và thành tiền ở `HDBong!E:E`; chương trình cũng ghi kết quả đã tính để Power BI đọc được ngay. `GhiChu!A` là công thức mảng động `SORT(UNIQUE(FILTER(SanLuongSanXuat[MatHang],…)))`, kèm kết quả đã tính; `GhiChu!C:F` là quy ước mã mặt hàng chép từ mẫu (`PRODUCT_CODE_NOTES`).

Thứ tự dòng cũng giống mẫu (Power BI không phụ thuộc thứ tự này): `HieuQuaSanXuat` xếp mọi kỳ lũy kế trước kỳ tháng; `VatTuPhuTung` theo công đoạn; `KetQuaSanXuatNhaMay` theo kỳ → nhà máy → chỉ tiêu; `ThongTinNhaMay` để các dòng tiêu hao nguyên liệu ở cuối; nước trong `PhanTichTieuThu` theo thứ tự dòng trên form như `CongNoPhaiThu`.

Các giá trị nghiệp vụ, kiểu cột và công thức cần dùng được tái tạo từ nguồn. Các trường hợp chỉnh tay có chủ đích phải được đưa vào cấu hình, không nhập trực tiếp vào file đích rồi mong chương trình tự giữ lại khi chạy lại.

## 7. Đơn vị, lũy kế và các ngoại lệ

- Tiền nguồn `Trđ` được nhân 1.000.000 để ra đồng; giá và lượng giữ đúng đơn vị của từng bảng.
- `Dien` dùng kWh; `HieuQuaSanXuat` giữ điện theo **Nghìn kWh** như mẫu.
- Phế liệu dùng kg và đồng/kg; sản lượng sợi, tồn kho, hợp đồng bông dùng tấn.
- Thực hiện lũy kế đọc trực tiếp từ nguồn. Kế hoạch lũy kế phụ thuộc loại chỉ tiêu: tổng, bình quân, bình quân gia quyền hoặc giá trị tháng cuối.
- Nguồn kế hoạch mỗi tháng có thể được sửa lại các tháng trước; lũy kế kế hoạch của Tn dùng dải T1…Tn trong **chính file Tn**.
- Riêng `PhanTichTieuThu`: doanh số lũy kế giữ cột lũy kế nguồn và ghi cảnh báo nếu khác cộng dồn.
- `TongHopKQBanHang`: cả lượng và tiền lũy kế đều đọc trực tiếp từ `1. KQ SXKD`; không cộng lại từ số tháng.
- Giữ số âm của khoản phế liệu thu hồi trong chi phí.
- Không thay mọi ô trống/lỗi thành 0. Quy tắc 0 áp dụng ở những bảng không phát sinh đã xác định.

Các điều chỉnh tái tạo mẫu được lưu ở `src/reporting/config/reference_overrides.json`:

1. **Giá bông kế hoạch T8/2026:** nguồn 18,4, mẫu 1,84. Chỉ đổi khi năm/tháng/chỉ tiêu và giá nguồn còn khớp cấu hình. Lũy kế 3,84 được giữ theo mẫu và có cảnh báo; chưa tự sửa thành 1,77 vì đó là thay đổi số liệu nghiệp vụ cần xác nhận riêng.
2. **Ngày hai hợp đồng T6/2026:** mẫu bổ sung 15/06/2026, nguồn chứa chữ Bông Mỹ/Bông Brazil. Chỉ áp dụng khi nguồn gốc, số lượng, đơn giá, năm/tháng và chữ nguồn khớp. Nhật ký ghi rõ ngày chưa được chứng từ nguồn xác nhận.
3. **Lỗi tồn kho T1/2026:** chỉ các ô `#REF!` đã được liệt kê mới được dùng 0, và phải có thành tiền tồn trên 3 tháng cùng hàng bằng 0. Các lỗi khác làm chương trình dừng.

Đặt `APPLY_REFERENCE_OVERRIDES=False` để bỏ hai điều chỉnh giá và ngày hợp đồng: khi đó T8 giữ 18,4 theo nguồn, ngày hợp đồng T6 để trống. Ngoại lệ tồn kho không phát sinh vẫn được kiểm soát riêng theo danh sách ô.

Một số cảnh báo khác được giữ để không che vấn đề nguồn:

- Lao động LK7T là 715 người, cao bất thường so với tháng báo cáo.
- Giá bông/xơ tổng SPH không có thì dùng bình quân đơn giản giá dương của nhà máy như quy trình đã học, không gọi là bình quân gia quyền.
- T1–T5 chưa có cột tồn trên chuyền nên dùng 0 theo mẫu.
- Mẫu gọi `Xuất khẩu tại chỗ` trong cột `QuocGia`; tên cột và nội dung được giữ đúng mẫu.
- `HDBong` (bảng `MuaBongTable`) không có cột loại giao dịch, nên chỉ xuất hợp đồng **Phát sinh trong kỳ**. Nếu nguồn có dòng **Ghi giảm kỳ trước**, dòng đó không xuất và được ghi cảnh báo `CONTRACT_REDUCTION_NOT_EXPORTED`; tổng hai phần vẫn được đối chiếu với nguồn.

Các cột nguồn không có trong mẫu thì không xuất và không đọc: giá nội địa/EXW/FOB của sợi (chỉ giữ CIF), đơn giá VNĐ tồn kho nguyên liệu, đơn giá phế liệu, đơn giá điện, sản lượng tiêu thụ theo nước. Cột xuất khác của tồn kho sợi vẫn được đọc để đối chiếu dòng tổng.

Đọc mô tả nghiệp vụ chi tiết trong `QUY_TAC_TONG_HOP_DuLieuBaoCao.md` nếu cần truy vết cách lấy từng cột.

### Sheet TongHopKQBanHang (bổ sung 19/09/2026)

Nguồn: file `KQ Tieu thu` từng tháng, sheet `1. KQ SXKD`. Xuất đúng 4 cột `Nam, Ky, ChiTieu, GiaTri`, bảng Excel `TongHopKQBanHangTable`.

- Hai chỉ tiêu lượng: Số lượng bán nội địa, Số lượng bán xuất khẩu. Đơn vị tấn, tối đa 6 chữ số thập phân, định dạng `#,##0.######`.
- Bảy chỉ tiêu tiền: Doanh thu nội địa, Doanh thu xuất khẩu, Giá vốn hàng bán, Chi phí lãi vay, Chi phí bán hàng, Chi phí quản lý doanh nghiệp, Lợi nhuận trước thuế. Nhân 1.000.000 và làm tròn về đồng nguyên, giữ số âm.
- Tìm cột bằng cả tiêu đề `THỰC HIỆN THÁNG BÁO CÁO` ở hàng 4 và đơn vị ở hàng 5, tránh cột ước thực hiện. Lượng ở I/J; tiền ở V/W (T1–T3), W/X (T4–T5), Y/Z (T6–T8).
- Mỗi chỉ tiêu xuất hai dòng Tn và LKnT, không thêm T12 năm trước cho sheet này. 8 tháng × 9 chỉ tiêu × 2 kỳ = 144 dòng.
- Kiểm tra kỳ, mã SPH, đơn vị, tổng lượng gồm nội địa + xuất khẩu + nội bộ. Thiếu chỉ tiêu/số liệu hoặc lỗi công thức sẽ dừng. Nội bộ chỉ dùng kiểm soát tổng, không xuất thành chỉ tiêu thứ 10.
- Lũy kế có điều chỉnh được giữ theo nguồn và ghi `SALES_SUMMARY_CUMULATIVE_ADJUSTMENT`. Ví dụ doanh thu nội địa LK4T là 32.118.491.933 đồng, không thay bằng tổng các tháng. Vị trí từng ô nằm trong `lineage` của nhật ký JSON.

### Sheet CongNoPhaiThu (bổ sung 24/09/2026)

Nguồn: file `KQ Tieu thu` từng tháng, sheet `2. Phai thu`. Bảng Excel `CongNoPhaiThuTable`, cột:

`Ky, QuocGia, KenhTieuThu, SanLuongBanHang, DoanhSoThangDong, DoanhSoLuyKeDong, TrongHan, QuaHanDuoi6T, QuaHan6Den12T, QuaHan12Den24T, QuaHan24Den36T, QuaHanTren36T`

- Đọc đủ 8 tháng rồi mới ghi. Mỗi tháng có đủ **cùng một danh sách**: `Việt Nam/Nội địa` (dòng nhóm Thị trường nội địa, gộp khách hàng nội bộ và khác), mọi nước `…/Xuất khẩu` có ít nhất một số khác 0 (sản lượng, doanh số, lũy kế, số dư hoặc tuổi nợ) ở **bất kỳ tháng nào** trong năm, và `Xuất khẩu tại chỗ/Xuất khẩu` (XKTC). Tên nước/kênh theo quy ước `PhanTichTieuThu`.
- Nước không phát sinh trong tháng vẫn có dòng giá trị 0. Ví dụ Nhật Bản bán lần đầu T3 vẫn có dòng T1, T2; Malaysia T2 chỉ có số dư mang sang từ năm trước, T3 bằng 0, T4 có bán. Dữ liệu 2026 T1–T8: 9 nước/kênh × 8 tháng = 72 dòng.
- Form bị thêm/bớt dòng nước giữa các tháng (Macau có đến T5, bỏ từ T6) được ghi `RECEIVABLE_COUNTRY_LIST_CHANGED`. Nếu form tháng sau không còn dòng của nước đã có lũy kế/số dư thì dừng; nếu tháng trước bằng 0 thì ghi 0 kèm `RECEIVABLE_COUNTRY_ROW_MISSING`.
- Tên nước sửa chính tả giữa các tháng được chuẩn hóa trong `COUNTRY_ALIASES` (`reporting/common.py`), dùng chung với `PhanTichTieuThu`: `Thổ Nhỹ Kỳ` (T1–T5) → `Thổ Nhĩ Kỳ` (cảnh báo `COUNTRY_NAME_ALIAS`). Thêm cặp mới vào đây khi form đổi tên nước.
- Dòng có số nhưng mất STT sẽ dừng, không bỏ sót.
- Không xuất dòng TỔNG CỘNG hoặc dòng nhóm xuất khẩu: cộng các dòng trong BI ra đúng tổng nguồn, không trùng; chương trình kiểm tra điều này cho từng tháng.
- Sản lượng: tấn, tối đa 6 chữ số thập phân. Tiền (doanh số, 6 nhóm tuổi nợ): nguồn triệu đồng × 1.000.000, làm tròn về **đồng nguyên** như các sheet khác.
- Tìm cột theo tiêu đề hàng `Diễn giải` và hàng phân loại tuổi nợ; đổi tiêu đề, thiếu số ở dòng tổng/nhóm hoặc trùng tên nước sẽ dừng.
- Dừng nếu: số dư công nợ ≠ tổng 6 nhóm tuổi nợ; tổng cộng ≠ nội địa + xuất khẩu; dòng nhóm ≠ tổng khách hàng/nước (+ XKTC) ở mọi cột. Nguồn chỉ chia nước ở cột Trong hạn, nên nếu có nợ quá hạn mà không chia theo nước, chương trình sẽ dừng thay vì gán sai nước.
- Chỉ cảnh báo: lệch với `1. KQ SXKD` (`RECEIVABLE_SALES_MISMATCH`); số dư công nợ bằng đúng doanh số tháng (`RECEIVABLE_EQUALS_MONTHLY_SALES`). Lũy kế doanh số có điều chỉnh đã được cảnh báo theo từng nước/kênh bằng `SALES_CUMULATIVE_ADJUSTMENT` (cùng ô nguồn với `PhanTichTieuThu`).

### Sheet TaiChinhNhaMay và tiêu hao nguyên liệu trong ThongTinNhaMay (bổ sung 24/09/2026)

Nguồn: file QTSX từng tháng, sheet `1. QTSX`, hai khối cột nhà máy `THÁNG BÁO CÁO` (NM1 = SPH-1, NM2 = SPH-2), mỗi khối gồm cột tháng và cột lũy kế.

- `TaiChinhNhaMay`: 4 chỉ tiêu `Doanh thu sản xuất`, `Lợi nhuận trước thuế`, `Nguyên giá TSCĐ`, `Giá trị còn lại TSCĐ`. Nguồn Trđ × 1.000.000, làm tròn về đồng nguyên bằng `dong()` như mọi sheet. 8 tháng × 2 kỳ × 2 nhà máy × 4 chỉ tiêu = 128 dòng.
- Nguyên giá và giá trị còn lại TSCĐ là **số dư cuối kỳ**; nguồn chỉ ghi ở cột lũy kế, nên Tn và LKnT dùng cùng số đó.
- `ThongTinNhaMay` có thêm `Bông tiêu hao` và `Xơ tiêu hao` (tấn, giữ số nguồn), xếp cuối sheet như mẫu: 2 × 2 nhà máy × 16 kỳ = 64 dòng.
- Kiểm tra đơn vị (`Trđ`, `Tấn`) và đối chiếu SPH-1 + SPH-2 với cột tổng SPH cho từng chỉ tiêu, từng kỳ.
- Khác biệt đã biết với mẫu: 8 ô giá trị còn lại TSCĐ có số nguồn đúng nửa đồng (ví dụ 26.567,3493565 Trđ). Mẫu làm tròn về số chẵn (26.567.349.356), chương trình làm tròn lên như các sheet khác (26.567.349.357).

Không cần cài thêm thư viện hoặc đổi cách chạy `main.py`. Đường dẫn đích trong README phản ánh cấu hình hiện tại bạn đã đặt: `DuLieuBaoCao_TongHop1.xlsx`. Để Power BI đọc trực tiếp kết quả, đặt `OUTPUT_FILE` là `D:\Dashboard\DuLieuBaoCao.xlsx` (đóng file trong Excel trước khi chạy; bản cũ được sao lưu vào `backups`).

## 8. Nhật ký và sao lưu

Sau mỗi lần chạy, thư mục cạnh file đích có:

```text
DataBase\
├── DuLieuBaoCao_TongHop1.xlsx
├── logs\
│   ├── DuLieuBaoCao_TongHop_<thời điểm>.log
│   └── DuLieuBaoCao_TongHop_<thời điểm>.json
└── backups\
    └── DuLieuBaoCao_TongHop_<thời điểm>.xlsx
```

- `.log`: dễ đọc bằng Notepad, gồm số dòng mỗi sheet, cảnh báo và lỗi.
- `.json`: chi tiết file nguồn, SHA-256, cấu hình, số liệu kiểm tra đối chiếu và vị trí nguồn của từng dòng xuất. Trong `lineage`, phần tử đầu tiên tương ứng hàng 2 của sheet đầu ra. Các hàng tính/gộp ghi vùng hoặc hàng nguồn liên quan.
- `backups`: chỉ có bản sao khi file đích đã tồn tại và có thể đọc được.

Quy trình ghi file: **tạo file tạm → đọc lại kiểm tra toàn bộ dữ liệu → sao lưu bản cũ → thay file đích**. Nếu lỗi trước bước thay thế, bản cũ còn nguyên. Nếu nguồn thay đổi trong lúc chạy, chương trình phát hiện qua hash và yêu cầu chạy lại. Không chỉnh sửa file nguồn.

Nếu không có quyền ghi ngay từ đầu thì có thể chưa tạo được nhật ký; khi đó đọc thông báo trên màn hình. Một số lỗi rất muộn như hết dung lượng lúc ghi nhật ký có thể xảy ra sau khi đã xuất workbook; kiểm tra thời điểm và nội dung file trước khi kết luận.

Không có chức năng tự xóa lịch sử. Có thể lưu trữ hoặc dọn `logs` và `backups` định kỳ theo nhu cầu của bạn.

## 9. Cấu trúc mã nguồn

```text
src/
├── main.py                         # Đường dẫn, tùy chọn, điểm chạy duy nhất
└── reporting/
    ├── models.py                   # Kiểu cấu hình, kết quả, kiểm tra/cảnh báo
    ├── sources.py                  # Tìm nguồn, đọc Excel, kiểm tra ngày/số
    ├── common.py                   # Đổi đơn vị, làm tròn, phân loại sản phẩm
    ├── production.py               # Kế hoạch, thực hiện, thông tin/tài chính nhà máy, chi phí
    ├── operations.py               # Sản lượng, vật tư phụ tùng, điện
    ├── inventory.py                # Tồn kho nguyên liệu và thành phẩm
    ├── sales.py                    # Giá sợi, phế liệu, HĐ bông, tiêu thụ
    ├── sales_summary.py            # 9 chỉ tiêu tổng hợp kết quả bán hàng
    ├── receivables.py              # Công nợ phải thu theo nước/kênh
    ├── validation.py               # Khóa trùng, danh mục và quan hệ dữ liệu
    ├── excel_writer.py             # Tạo Excel, tên bảng, định dạng, GhiChu, kiểm tra bản lưu
    ├── pipeline.py                 # Gọi các bước theo đúng trình tự, sắp thứ tự dòng
    └── config/
        ├── schemas.json            # Tên sheet, thứ tự cột đúng mẫu Dashboard
        ├── product_names.json      # Ánh xạ tên mặt hàng đã kiểm chứng
        └── reference_overrides.json# Các ngoại lệ riêng của mẫu
src/tests/
└── test_reporting.py               # Kiểm thử đơn vị và đối chiếu tích hợp
src/requirements.txt
src/README.md
```

Luồng đọc dễ theo dõi: `main.py` → `pipeline.run_report()` → `sources.discover()` → các mô-đun nghiệp vụ → `validation.validate()` → `excel_writer.write_report()`.

Không có tác vụ tự chạy khi import mô-đun. Có thể tái sử dụng từ ứng dụng Python khác bằng `Settings` và `run_report`, nhưng người dùng thông thường chỉ cần `main.py`.

## 10. Kiểm thử

Không cần cài thêm thư viện. Tại thư mục `G:\DOAN`:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s src/tests -v
```

Bộ kiểm thử kiểm tra: thiếu tháng, trùng file, nhiều năm, file khóa Excel, khóa dữ liệu trùng, lỗi Excel, công thức chưa được tính/lưu, số âm, phân loại sản phẩm, chuỗi bắt đầu bằng dấu `=`, giữ file cũ khi lỗi ghi và dữ liệu công thức đã lưu.

Kiểm thử tích hợp dùng bộ nguồn T1–T8/2026 tại `MonthlyReportingData` và file đối chứng `D:\Dashboard\DuLieuBaoCao.xlsx`. Nếu một trong hai không có, phần kiểm thử tích hợp được bỏ qua (`skipped`), các kiểm thử còn lại vẫn chạy. Có thể chỉ định mẫu đối chứng khác **cùng bộ T1–T8/2026**:

```powershell
$env:REPORT_REFERENCE_FILE = "D:\Mau\DuLieuBaoCao.xlsx"
.\.venv\Scripts\python.exe -m unittest discover -s src/tests -v
```

Phần tích hợp là kiểm thử hồi quy cho bộ T1–T8 đã học; khi bạn thêm T9 hoặc sửa số liệu nguồn, cần cập nhật bộ đối chứng/kiểm thử tích hợp tương ứng. Chương trình chạy chính không giới hạn ở T8.

Kết quả đã kiểm tra trên bộ hiện tại: **32 bài kiểm thử đạt**. Toàn bộ **19 sheet dữ liệu / 2.822 dòng** khớp mẫu theo tiêu đề, **thứ tự dòng** và tên bảng Excel, sau chuẩn hóa sai số số thực tới 6 chữ số thập phân; riêng 8 ô nửa đồng của `TaiChinhNhaMay` lệch 1 đồng như nêu ở mục 7. Danh sách mặt hàng ở `GhiChu` khớp mẫu. Cột phân loại và thành tiền có cả công thức lẫn giá trị đã lưu. Chưa kiểm tra việc Excel tự tính lại sau khi người dùng sửa trực tiếp ô đầu ra; dữ liệu được kiểm chứng bằng tính toán Python, kiểm soát nguồn và đọc lại workbook đã lưu.

## 11. Xử lý lỗi thường gặp

| Thông báo/tình huống | Cách xử lý |
|---|---|
| `py`/`python` không được nhận diện | Cài Python, bật PATH và mở lại Terminal; hoặc dùng đường dẫn đầy đủ tới python.exe |
| `No module named openpyxl`/`xlsxwriter` | Chạy lệnh cài requirements bằng đúng Python dùng chạy main |
| Lỗi mạng khi `pip install` | Kiểm tra Internet/proxy của đơn vị; cài thư viện trước khi chạy offline |
| Không thấy thư mục nguồn | Kiểm tra `SOURCE_FOLDER`, tên ổ đĩa và quyền đọc |
| Không tìm thấy đúng một file tháng | Kiểm tra tên file chứa `QTSX`/`KQ Tieu thu`, chuyển bản trùng ra ngoài |
| Thiếu tháng | Bổ sung đủ T1…tháng cuối hoặc chuyển thư mục tháng vượt khoảng ra ngoài nguồn |
| Nhiều năm | Đặt `REPORT_YEAR` cụ thể |
| Ngày báo cáo khác thư mục | Kiểm tra ngày trong workbook và tên `BCQT_MMYYYY` |
| Công thức chưa có giá trị đã lưu | Mở nguồn trong Excel, bật tính tự động, tính lại, lưu rồi chạy lại |
| Lỗi `#REF!`, `#DIV/0!` chưa được chấp nhận | Kiểm tra đúng ô nguồn được báo; không tự thay hàng loạt bằng 0 |
| Không ghi được file đích | Đóng Excel/ứng dụng khóa file; kiểm tra quyền ghi; hoặc đổi `OUTPUT_FILE` |
| Đối chiếu tổng không khớp | Đọc file/sheet/ô, số thực tế và số kiểm soát trong thông báo/log; sửa nguyên nhân nguồn hoặc cập nhật ánh xạ |
| Bố cục/đơn vị thay đổi | Kiểm tra mẫu báo cáo tháng mới và cập nhật mô-đun tương ứng |

Để cài trên máy không có Internet, dùng một máy có Internet chạy:

```powershell
py -m pip download -r src/requirements.txt -d wheels
```

Chép thư mục `wheels` và toàn bộ thư mục `src` sang máy đích, sau đó:

```powershell
.\.venv\Scripts\python.exe -m pip install --no-index --find-links wheels -r src/requirements.txt
```

Lệnh chạy cuối cùng vẫn chỉ là:

```powershell
.\.venv\Scripts\python.exe .\src\main.py
```
