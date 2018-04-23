RD /Q /S "%CD%\build\"
RD /Q /S "%CD%\dist\"
pyinstaller.exe --onefile --icon=icon.ico --noconsole main.py
