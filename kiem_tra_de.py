"""Kiểm tra một đề tự soạn trước khi đưa vào app.

    py kiem_tra_de.py de_thi\\Excel_Mau      kiểm tra đề
    py kiem_tra_de.py --luat                 xem danh sách luật chấm

Công cụ sẽ:
  1. Báo lỗi khai báo trong de.json (sai tên luật, thiếu tham số, thiếu file).
  2. Chấm file gốc – mọi nhiệm vụ phải SAI (nếu ĐÚNG sẵn thì luật chấm quá dễ).
  3. Nếu có thư mục dap_an/ chứa file đã làm đúng – mọi nhiệm vụ phải ĐÚNG.
"""
import sys
from pathlib import Path

from mos import rules
from mos.custom import check_exam, load_exam


def main(argv) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if argv[0] == "--luat":
        print(rules.describe())
        return 0

    folder = Path(argv[0])
    errors, rows = check_exam(folder)
    if errors:
        print("LỖI KHAI BÁO:")
        for e in errors:
            print("  ✗", e)
        return 1
    exam, _ = load_exam(folder)
    print(f"Đề: {exam.name}  ({exam.code})\n")

    problems, current = 0, None
    for r in rows:
        if r["project"] != current:
            current = r["project"]
            print(f"\n■ {current}  [{r['file']}]")
        line = f"  {r['index']:>2}. {r['task'][:70]:<70}  file gốc: {'ĐÚNG ⚠' if r['start_ok'] else 'sai ✓'}"
        problems += r["start_ok"]
        if r["answer_ok"] is None:
            line += "   (chưa có đáp án)"
        else:
            line += f"   đáp án: {'đúng ✓' if r['answer_ok'] else 'SAI ✗'}"
            problems += not r["answer_ok"]
            if r["error"]:
                line += f"  ({r['error']})"
        print(line)

    print("\nKẾT QUẢ:", "Đề ổn ✓" if not problems else f"Có {problems} điểm cần xem lại (⚠ / ✗ ở trên)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
