from tkinter import Tk, filedialog
from core.monthly_excel import monthly_to_excel

# Tkinter 창 숨기기
root = Tk()
root.withdraw()

# Excel 파일 선택
excel_file = filedialog.askopenfilename(
    title="Excel 파일 선택",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
)

# 파일 선택 취소 시
if not excel_file:
    print("파일 선택이 취소되었습니다.")
else:
    monthly_to_excel(excel_file)