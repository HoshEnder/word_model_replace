@echo off

REM Get the path of the folder containing the script
set "script_folder=%~dp0"

REM Define the relative path of the template from the script  
set "template=%script_folder%normal.dotm"

REM Define the user's Templates folder
set "templates_folder=%appdata%\Microsoft\Templates"

REM Initialize the counter
set counter=0

REM Increment the counter until we find an available filename
:loop
if exist "%templates_folder%\Normal.dotm_%counter%" (
  set /a counter+=1
  goto loop
)

REM Rename the old Normal.dotm template with the counter
if exist "%templates_folder%\Normal.dotm" (
  echo Renaming the old Normal.dotm template to Normal.dotm_%counter%...
  ren "%templates_folder%\Normal.dotm" "Normal.dotm_%counter%"
)

echo Copying the custom Word template...
xcopy /y "%template%" "%templates_folder%\"

echo End of script.
pause
