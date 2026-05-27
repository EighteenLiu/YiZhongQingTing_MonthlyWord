from docx import Document
from docx.shared import Cm

input_path = r"D:\桌面\西城区2026年5月可回收物交投点检查通报.docx"
output_path = r"D:\pycharm\YiZhongQingTing_MonthlyWord\monthly_report_system\西城区2026年5月可回收物交投点检查通报_图片统一宽度.docx"

doc = Document(input_path)

count = 0
for shape in doc.inline_shapes:
    shape.width = Cm(5.0)
    shape.height = Cm(3.0)
    count += 1

doc.save(output_path)

print(f"处理完成，共调整 {count} 张图片")
print(f"输出文件：{output_path}")