@echo off
pyinstaller --noconsole --ico="C:\python\extracao_itens-sei\icon\download_sei.ico" --noconfirm --onedir ^
--add-data "C:\python\extracao_itens-sei\excel\;excel" ^
--add-data "C:\python\extracao_itens-sei\chromedriver-win64\*;chromedriver-win64" ^
--add-data "C:\python\extracao_itens-sei\login_sei.py;." ^
--add-data "C:\python\extracao_itens-sei\README.pdf;." ^
extracao_itens-sei.py
pause
