@echo off

rem call venv\Scripts\activate
call py -m pipenv shell
rem pyInstaller --onefile --noconsole csvExtractor.py
pyInstaller csvExtractor.spec
rem move dist\csvExtractor.exe dist\csvExtractor.exe

pause