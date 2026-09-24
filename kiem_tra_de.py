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
from mos.core import run_check
from mos.custom import load_exam


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
    exam, errors = load_exam(folder)
    if errors:
        print("LỖI KHAI BÁO:")
        for e in errors:
            print("  ✗", e)
        return 1
    print(f"Đề: {exam.name}  ({exam.code}, {exam.minutes} phút)\n")

    problems = 0
    for project in exam.projects:
        print(f"■ {project.name}  [{project.filename}]")
        answer = folder / "dap_an" / project.filename
        for i, task in enumerate(project.tasks, start=1):
            start_ok, start_err = run_check(task, folder / project.filename)
            line = f"  {i:>2}. {task.title[:70]:<70}  file gốc: {'ĐÚNG ⚠' if start_ok else 'sai ✓'}"
            if start_ok:
                problems += 1
            if answer.is_file():
                ok, err = run_check(task, answer)
                line += f"   đáp án: {'đúng ✓' if ok else 'SAI ✗'}"
                if not ok:
                    problems += 1
                    if err:
                        line += f"  ({err})"
            print(line)
        if not answer.is_file():
            print(f"     (chưa có dap_an/{project.filename} để kiểm tra chiều đúng)")
        print()

    print("KẾT QUẢ:", "Đề ổn ✓" if not problems else f"Có {problems} điểm cần xem lại (⚠ / ✗ ở trên)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
