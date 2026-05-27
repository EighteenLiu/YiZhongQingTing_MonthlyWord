# 可回收物交投点月报生成系统

本项目是一个本地桌面窗口工具，用于选择 Excel 检查数据和 Word 月报模板，自动提取检查记录和现场图片，统计问题指标，并生成 Word 月报 `.docx`。

## 运行

双击或执行：

```bat
run_app.bat
```

首次运行会在项目目录下创建 `.venv` 虚拟环境并安装依赖；后续会直接复用该环境启动窗口。

## 使用流程

1. 选择 Excel 检查数据文件，支持 `.xls` / `.xlsx`。
2. 选择 Word 月报模板，支持 `.docx`。
3. 填写报告标题、检查日期、报告月份、输出目录和输出文件名。
4. 点击“生成 Word 月报”。
5. 生成完成后可在窗口中查看统计、预览、日志，并点击“打开文件”查看 Word。

## 输出与临时文件

- 生成的 Word 默认保存到 `monthly_report_system/output/`。
- 处理过程中产生的临时文件保存到 `monthly_report_system/temp/`，单次生成结束后会自动清理。
- `.venv`、`temp`、`output`、缓存文件均已通过 `.gitignore` 忽略。

## 开发验证

```bat
python -m pytest -q
```
