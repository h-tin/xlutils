python -m venv .venv
call .\.venv\Scripts\activate.bat

pip install pyinstaller
pip install -r requirements.txt
pyinstaller --onefile %~1

call .\.venv\Scripts\deactivate.bat

move .\dist\*.exe .\

rmdir .\.venv /S /Q
rmdir .\build /S /Q
rmdir .\dist /S /Q
del %~n1.spec
