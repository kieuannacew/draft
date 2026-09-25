"""Nhập bộ đề có sẵn (thư mục hoặc file .zip) vào app.

    py nhap_de.py D:\\MOS_Word365_DeThucTe          nhập mọi đề De_xx trong thư mục
    py nhap_de.py D:\\MOS_Word365_DeThucTe.zip      nhập thẳng từ file nén

Mỗi đề De_xx thành một bài thi riêng trên Trang chủ; câu Word được chấm tự động.
Đề được lưu vào thư mục dữ liệu của app (~/MOS_Practice/de_thi hoặc thư mục trong cau_hinh.json).
"""
import sys
from pathlib import Path

from mos import nhap_de


def main(argv) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    try:
        done, warnings = nhap_de.import_path(Path(argv[0]))
    except nhap_de.ImportError_ as exc:
        print("LỖI:", exc)
        return 1
    for folder in done:
        print("  ✓", folder)
    for w in warnings:
        print("  ⚠", w)
    print(f"\nĐã nhập {len(done)} đề. Mở app để làm bài.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
