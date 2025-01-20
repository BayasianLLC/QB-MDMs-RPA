@echo off
echo Installing MDM Processor...

echo Creating directories...
mkdir "%PROGRAMFILES%\Wesco\MDM Processor" 2>nul
mkdir "%PROGRAMFILES%\Wesco\MDM Processor\config" 2>nul
mkdir "%PROGRAMFILES%\Wesco\MDM Processor\logs" 2>nul

echo Copying files...
xcopy /Y /E /I "*.exe" "%PROGRAMFILES%\Wesco\MDM Processor\"
xcopy /Y /E /I "config\*" "%PROGRAMFILES%\Wesco\MDM Processor\config\"

echo Creating shortcuts...
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Desktop\MDM Processor.lnk');$s.TargetPath='%PROGRAMFILES%\Wesco\MDM Processor\MDM_Processor.exe';$s.Save()"
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\MDM Processor.lnk');$s.TargetPath='%PROGRAMFILES%\Wesco\MDM Processor\MDM_Processor.exe';$s.Save()"

echo Installation complete!
pause
