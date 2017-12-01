REM
REM Create temporary directory
REM

mkdir C:\Temp
cd /D C:\Temp

REM
REM Make sure Network Discovery is allowed through the firewall so we can resolve the hostname
REM
netsh advfirewall firewall set rule group="network discovery" new enable=yes

REM Check to only add it if we haven't already added it
netsh advfirewall firewall show rule name="ICMP Allow incoming V4 echo request" >nul
if ERRORLEVEL 1 netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
netsh advfirewall firewall show rule name="ICMP Allow incoming V6 echo request" >nul
if ERRORLEVEL 1 netsh advfirewall firewall add rule name="ICMP Allow incoming V6 echo request" protocol=icmpv6:8,any dir=in action=allow

REM If you did add a bunch, delete them all with the following and then run the above
REM netsh advfirewall firewall delete rule name="ICMP Allow incoming V4 echo request"
REM netsh advfirewall firewall delete rule name="ICMP Allow incoming V6 echo request"

REM
REM Auto login student
REM
If Not %computername%==TEACHER (
REG ADD "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f
REG ADD "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d Student /f
REG ADD "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d "" /f
)

REM
REM Disable auto-updates this week while we're programming
REM
sc stop wuauserv
sc config wuauserv start=disabled

REM
REM Remove the annoying windows that auto run for no reason
REM
rmdir /S /Q "C:\Users\sunset\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Google"
rmdir /S /Q "C:\Users\sunset\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Microsoft.BingWeather_8wekyb3d8bbwe"
rmdir /S /Q "C:\Users\default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Google"
rmdir /S /Q "C:\Users\default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Microsoft.BingWeather_8wekyb3d8bbwe"
rmdir /S /Q "C:\Users\student\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Google"
rmdir /S /Q "C:\Users\student\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Microsoft.BingWeather_8wekyb3d8bbwe"

REM
REM Distribute files
REM

robocopy "\\teacher\Server\Files" "C:\Users\Student\Documents" /E
rmdir /S /Q "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Defender"
rmdir /S /Q "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Football"
rmdir /S /Q "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Lander"
rmdir /S /Q "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Maze"
rmdir /S /Q "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Roguelike"
del "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Defender.meta"
del "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Football.meta"
del "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Lander.meta"
del "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Maze.meta"
del "C:\Users\Student\Documents\Game Programming\PlaygroundProject\PlaygroundProject\Assets\Scenes\Roguelike.meta"
REM /purge
copy "\\teacher\Server\Unity Documentation.lnk" C:\Users\Public\Desktop\

REM
REM Install software if it doesn't exist
REM

REM If Not Exist "C:\Program Files (x86)\K-Lite Codec Pack" (
REM copy "\\teacher\Server\K-Lite_Codec_Pack_1324_Mega.exe" C:\Temp\
REM start /wait K-Lite_Codec_Pack_1324_Mega.exe /verysilent /suppressmsgboxes /norestart /sp-
REM )

If Not Exist "C:\Program Files (x86)\obs-studio" (
copy "\\teacher\Server\OBS-Studio-19.0.3-Full-Installer.exe" C:\Temp\
start /wait OBS-Studio-19.0.3-Full-Installer.exe /S
)

If Not Exist "C:\Program Files (x86)\Scratch" (
copy "\\teacher\Server\Scratch\Scratch1.4.msi" C:\Temp\
copy "\\teacher\Server\Scratch\Scratch.ini" C:\Temp\
msiexec /i Scratch1.4.msi /qn /l* C:\Temp\scratch.log.txt
)

If Not Exist "C:\Program Files\Unity" (
copy "\\teacher\Server\UnitySetup64-5.6.1f1.exe" C:\Temp\
start /wait UnitySetup64-5.6.1f1.exe /s
)

If Not Exist "C:\Program Files (x86)\LEGO Software\LEGO MINDSTORMS NXT" (
robocopy "\\teacher\Server\MINDSTORMS NXT ISO" C:\Temp\ /E
start /wait "MINDSTORMS NXT ISO\setup.exe" /q /AcceptLicenses yes /r:n
)

If Not Exist "C:\Program Files (x86)\LEGO Software\LEGO MINDSTORMS EV3 Home Edition" (
copy "\\teacher\Server\LMS-EV3-WIN32-ENUS-01-02-02-full-setup.exe" C:\Temp\
copy "\\teacher\Server\LMS-EV3-WIN32-ENUS-01-02-02-full-setup-specfile.txt" C:\Temp\
start /wait LMS-EV3-WIN32-ENUS-01-02-02-full-setup.exe C:\Temp\LMS-EV3-WIN32-ENUS-01-02-02-full-setup-specfile.txt /q /AcceptLicenses yes /r:n
)

If Not Exist "C:\Program Files (x86)\Microsoft Visual Studio Tools for Unity" (
copy "\\teacher\Server\Visual Studio 2015 Unity Plugin - vstu2015.msi" C:\Temp\
msiexec /i "Visual Studio 2015 Unity Plugin - vstu2015.msi" /qn /l* C:\Temp\vstu2015.log.txt
)

If Not Exist "C:\Program Files\LibreOffice 5" (
copy \\teacher\Server\LibreOffice_5.3.4_Win_x64.msi C:\Temp\
msiexec /i LibreOffice_5.3.4_Win_x64.msi /qn /l* C:\Temp\libreoffice.log.txt
)

If Not Exist "C:\Program Files\Blender Foundation\Blender\2.76" (
copy \\teacher\Server\blender-2.76b-windows64.msi C:\Temp\
msiexec /i blender-2.76b-windows64.msi /qn /l* C:\Temp\blender.log.txt
copy \\teacher\Server\Blender.lnk C:\Users\Public\Desktop\
)

If Not Exist "C:\Program Files (x86)\Audacity\audacity.exe" (
copy \\teacher\Server\audacity-win-2.1.2.exe C:\Temp\
start /wait audacity-win-2.1.2.exe /verysilent /suppressmsgboxes /norestart /sp-
)

If Not Exist "C:\Program Files (x86)\FFmpeg for Audacity" (
copy \\teacher\Server\ffmpeg-win-2.2.2.exe C:\Temp\
start /wait ffmpeg-win-2.2.2.exe /verysilent /suppressmsgboxes /norestart /sp-
)

If Not Exist "C:\Program Files (x86)\Audacity\Plug-Ins\FFTW_docs" (
copy \\teacher\Server\LADSPA_plugins-win-0.4.15.exe C:\Temp\
start /wait LADSPA_plugins-win-0.4.15.exe /verysilent
)

If Not Exist "C:\Program Files (x86)\Lame For Audacity" (
copy \\teacher\Server\Lame_v3.99.3_for_Windows.exe C:\Temp\
start /wait Lame_v3.99.3_for_Windows.exe /verysilent
)

If Not Exist "C:\Anaconda3" (
copy \\teacher\Server\Anaconda3-4.1.1-Windows-x86_64.exe C:\Temp\
start /wait Anaconda3-4.1.1-Windows-x86_64.exe /S /InstallationType=AllUsers /AddToPath=1 /RegisterPython=1 /NoRegistry=0 /D=C:\Anaconda3
)

If Not Exist "C:\Program Files\GIMP 2" (
copy \\teacher\Server\gimp-2.8.18-setup.exe C:\Temp\
gimp-2.8.18-setup.exe /VERYSILENT /NORESTART
)

If Not Exist "C:\Program Files (x86)\Notepad++" (
copy \\teacher\Server\npp.6.9.2.Installer.exe C:\Temp\
npp.6.9.2.Installer.exe /S
)

If Not Exist "C:\Program Files\Programmer's Notepad" (
copy \\teacher\Server\pn2402378_multilang.exe C:\Temp\
pn2402378_multilang.exe /SILENT /DIR="%ProgramFiles%\Programmer's Notepad"
)

If Not Exist "C:\Program Files (x86)\Microsoft VS Code" (
copy \\teacher\Server\VSCodeSetup-stable.exe C:\Temp\
VSCodeSetup-stable.exe /VERYSILENT /NORESTART
)

If Not Exist "C:\Program Files (x86)\Microsoft Visual Studio 14.0" (
robocopy "\\teacher\Server\Community 2015" C:\Temp\ /E
start /wait vs_community.exe /AdminFile "C:\Temp\AdminDeployment.xml" /Quiet /NoWeb /NoRestart
)

REM
REM Remove temporary directory
REM
cd C:\
rmdir /S /Q C:\Temp
