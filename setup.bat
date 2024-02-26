ECHO OFF
::Salvam directoriul in care se gaseste programul exe
set current_dir=%cd%

:: Chemam python si gasim directoriul in care este instalat si salvam temporar acest path in fisierul text.txt
python -c "import sys;print(sys.executable)" >> text.txt
:: Asociem variabila python_path
set /p python_path= < text.txt
del text.txt

:: Acum avem variabila python_path, cu care vom lucra
:: Vom merge cu doua directorii mai sus, de exemplu, python_path este C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe
:: Avem nevoie sa mergem in Python312

for /D %%I in (%python_path%) do (   
SET parent_dir= %%~dpI
)

:: Acum variabila parent_dir va fi egala cu C:\Users\user\AppData\Local\Programs\Python\Python312\

:: Vom crea aici un directoriu numit MyScripts

if not exist "%parent_dir%\MyScripts" mkdir %parent_dir%\MyScripts

:: Schimbam directoriul curent in  \MyScripts

chdir %parent_dir%\MyScripts

:: facem git clone
git clone --branch main https://github.com/Johnny-Sanduleac/Evidenta_grupelor.git

:: Acum avem directoriul Evidenta_grupelor
chdir %parent_dir%\MyScripts\Evidenta_grupelor

:: instalam requirements.txt
pip install -r requirements.txt

:: Acum, ca fisierele excel sunt pregatite, le vom copia in mapa unde este exe-ul
mkdir "%current_dir%\Evidenta_grupelor"

:: Revenim la MyScripts
chdir %parent_dir%\MyScripts\Evidenta_grupelor
:: Cautam prin toate fisierele din mapa %parent_dir%\MyScripts\Evidenta_grupelor si copiem excel-urile in %current_dir%\Evidenta_grupelor

for %%a in (*.xlsx) do (xcopy /s %parent_dir%\MyScripts\Evidenta_grupelor\%%a  "%current_dir%\Evidenta_grupelor"
echo %%a)

:: Vom crea in acest directoriu si un scenariu pentru rularea script-ului main_script
chdir "%current_dir%\Evidenta_grupelor"

IF EXIST email_agent.bat DEL /F email_agent.bat
ECHO ECHO OFF >> email_agent.bat
ECHO chdir %parent_dir%\MyScripts\Evidenta_grupelor  >> email_agent.bat
ECHO python main_script.py >> email_agent.bat
ECHO ECHO "SUCCESS" >> email_agent.bat

ECHO "Process finsihed with success!"

PAUSE