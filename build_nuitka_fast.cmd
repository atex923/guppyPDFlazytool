@echo off
setlocal

python -m nuitka ^
  --standalone ^
  --windows-console-mode=disable ^
  --enable-plugin=tk-inter ^
  --nofollow-import-to=tkinterdnd2 ^
  --nofollow-import-to=paddleocr ^
  --nofollow-import-to=paddle ^
  --nofollow-import-to=easyocr ^
  --nofollow-import-to=rapidocr_onnxruntime ^
  --nofollow-import-to=onnxruntime ^
  --nofollow-import-to=cv2 ^
  --nofollow-import-to=shapely ^
  --nofollow-import-to=skimage ^
  --nofollow-import-to=scipy ^
  --nofollow-import-to=matplotlib ^
  --nofollow-import-to=torch ^
  --nofollow-import-to=tensorflow ^
  --output-dir=build ^
  --remove-output ^
  Guppy_PDFlazyTool.py

endlocal
