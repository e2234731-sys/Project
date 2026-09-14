@echo off
chcp 65001 >nul
echo ========================================================
echo   校准报告智能归档系统 - 全量独立单文件 EXE 构建程序
echo ========================================================
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
    --name="校准报告智能归档系统" ^
    --collect-all pymupdf ^
    --collect-all fitz ^
    --collect-all rapidocr_onnxruntime ^
    --collect-all pyclipper ^
    --collect-all shapely ^
    --collect-all yaml ^
    --collect-all onnxruntime ^
    --hidden-import=pymupdf ^
    --hidden-import=pymupdf.mupdf ^
    --hidden-import=fitz ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v2_cls ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v3_det ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v3_rec ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v2_cls.text_cls ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v3_det.text_detect ^
    --hidden-import=rapidocr_onnxruntime.ch_ppocr_v3_rec.text_recognize ^
    --hidden-import=rapidocr_onnxruntime.utils ^
    --hidden-import=pyclipper ^
    --hidden-import=shapely ^
    --hidden-import=yaml ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=PyQt5 ^
    main_gui.py

echo.
if exist "dist\校准报告智能归档系统.exe" (
    copy /y "dist\校准报告智能归档系统.exe" ".\校准报告智能归档系统.exe"
    echo ========================================================
    echo   [SUCCESS] 单文件可执行程序已生成: .\校准报告智能归档系统.exe
    echo ========================================================
) else (
    echo [ERROR] 打包失败，请检查上方日志。
)
