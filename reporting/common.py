"""Các phép đổi đơn vị và chuẩn hóa văn bản có phạm vi rõ ràng."""
import math
import re
import unicodedata


def text(value) -> str:
    return ' '.join(str(value if value is not None else '').split())


def key(value) -> str:
    return unicodedata.normalize('NFC', text(value)).casefold()


# Tên nước bị sửa chính tả giữa các tháng (T1–T5 'Thổ Nhỹ Kỳ', từ T6 'Thổ Nhĩ Kỳ').
# Chuẩn hóa để một nước không bị tách thành hai dòng khi ghép cả năm.
COUNTRY_ALIASES = {'thổ nhỹ kỳ': 'Thổ Nhĩ Kỳ'}


def country_name(value) -> str:
    name = unicodedata.normalize('NFC', text(value))
    return COUNTRY_ALIASES.get(key(name), name)


def rounded(value: float, digits: int=0):
    """Ghi chú triển khai:
- Giữ cách Math.round của quy trình cũ, kể cả số âm; tránh banker rounding."""
    factor = 10 ** digits
    result = math.floor(value * factor + 0.5) / factor
    return int(result) if digits == 0 else result


def dong(million: float) -> int:
    return rounded(million * 1000000)


def normalize_product(name: str, mapping: dict[str, str]) -> str:
    """Ghi chú triển khai:
- Chỉ sửa khoảng trắng/hậu tố wax cho tên mới, không tự đổi thành phần sợi."""
    cleaned = text(name)
    if cleaned in mapping:
        return mapping[cleaned]
    return re.sub('\\s*-?\\s*[wW]$', ' W', cleaned)


def classify_product(name: str) -> tuple[str, str, str, int | None]:
    """Tương đương 4 công thức C:F của mẫu, gồm phân biệt Cm với Cd."""
    upper = name.upper()
    material = 'CM/Modal' if 'CM/MODAL' in upper else 'Recycle' if 'R)' in upper else 'Thường'
    yarn = next((label for token, label in [('CM/MODAL', 'Sợi CM/Modal'), ('CVC', 'Sợi CVC'), ('TC', 'Sợi TC')] if token in upper), '')
    combing = 'Sợi chải kỹ' if 'CM/MODAL' in upper or 'Cm' in name else 'Sợi chải thô' if 'Cd' in name else ''
    count = int(name.split('/')[0]) if re.match('^\\d+/', name) else None
    return (material, yarn, combing, count)
