REM
REM Create temporary directory
REM

mkdir C:\Temp
cd /D C:\Temp

REM
REM Update mandatory profile
REM

robocopy \\server.yapsda.org\Profiles\student.V2 C:\Users\defaultstudent.V2 /MIR /SEC
icacls C:\Users\defaultstudent.V2 /T /setowner "yapsda.org\student"

REM
REM Make sure users don't change public documents
REM

robocopy \\server.yapsda.org\Software\Public C:\Users\Public /MIR /SEC

REM
REM Install software if it doesn't exist
REM

REM start /wait lmms-1.1.3-win64.exe /s

If Not Exist "C:\Program Files (x86)\Broderbund" (
copy "\\server.yapsda.org\Software\Mavis Beacon.lnk" C:\Users\Public\Desktop\
mkdir "C:\Program Files (x86)\Broderbund"
robocopy "\\server.yapsda.org\Software\Broderbund" "C:\Program Files (x86)\Broderbund" /E
)

REM If Not Exist "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" (
REM copy \\server.yapsda.org\Software\vlc-2.2.4-win32.exe C:\Temp\
REM start /wait vlc-2.2.4-win32.exe /L=1033 /S
REM )

If Exist "C:\Program Files\Blender Foundation\Blender\2.77" (
REM copy \\server.yapsda.org\Software\blender-2.77-windows64.msi C:\Temp\
REM msiexec /q /x blender-2.77-windows64.msi
rmdir /S /Q "C:\Program Files\Blender Foundation"
)

If Not Exist "C:\Program Files\Mozilla Firefox" (
copy \\server.yapsda.org\Software\FirefoxSetup51.0.1.exe C:\Temp\
start /wait FirefoxSetup51.0.1.exe -ms
)

If Not Exist "C:\Program Files (x86)\Blender Foundation\Blender\2.76" (
copy \\server.yapsda.org\Software\blender-2.76b-windows32.msi C:\Temp\
msiexec /qn /l* C:\Temp\blender.log.txt /a blender-2.76b-windows32.msi
move "C:\Blender Foundation" "C:\Program Files (x86)"
del "C:\blender-2.76b-windows32.msi"
copy \\server.yapsda.org\Software\Blender.lnk C:\Users\Public\Desktop\
)

If Not Exist "C:\Program Files (x86)\Audacity\audacity.exe" (
copy \\server.yapsda.org\Software\audacity-win-2.1.2.exe C:\Temp\
start /wait audacity-win-2.1.2.exe /verysilent /suppressmsgboxes /norestart /sp-
)

If Not Exist "C:\Program Files (x86)\FFmpeg for Audacity" (
copy \\server.yapsda.org\Software\ffmpeg-win-2.2.2.exe C:\Temp\
start /wait ffmpeg-win-2.2.2.exe /verysilent /suppressmsgboxes /norestart /sp-
)

If Not Exist "C:\Program Files (x86)\Audacity\Plug-Ins\FFTW_docs" (
copy \\server.yapsda.org\Software\LADSPA_plugins-win-0.4.15.exe C:\Temp\
start /wait LADSPA_plugins-win-0.4.15.exe /verysilent
)

If Not Exist "C:\Program Files (x86)\Lame For Audacity" (
copy \\server.yapsda.org\Software\Lame_v3.99.3_for_Windows.exe C:\Temp\
start /wait Lame_v3.99.3_for_Windows.exe /verysilent
)

If Not Exist "C:\Program Files\SumatraPDF\SumatraPDF.exe" (
copy \\server.yapsda.org\Software\SumatraPDF-3.1.2-64-install.exe C:\Temp\
start /wait SumatraPDF-3.1.2-64-install.exe /s /register
)

If Not Exist "C:\Anaconda3" (
copy \\server.yapsda.org\Software\Anaconda3-4.1.1-Windows-x86_64.exe C:\Temp\
start /wait Anaconda3-4.1.1-Windows-x86_64.exe /S /InstallationType=AllUsers /AddToPath=1 /RegisterPython=1 /NoRegistry=0 /D=C:\Anaconda3
)

If Not Exist "C:\Program Files\GIMP 2" (
copy \\server.yapsda.org\Software\gimp-2.8.18-setup.exe C:\Temp\
gimp-2.8.18-setup.exe /VERYSILENT /NORESTART
)

If Not Exist "C:\Program Files (x86)\Notepad++" (
copy \\server.yapsda.org\Software\npp.6.9.2.Installer.exe C:\Temp\
npp.6.9.2.Installer.exe /S
)

If Not Exist "C:\Program Files\Programmer's Notepad" (
copy \\server.yapsda.org\Software\pn2402378_multilang.exe C:\Temp\
pn2402378_multilang.exe /SILENT /DIR="%ProgramFiles%\Programmer's Notepad"
)

If Not Exist "C:\Program Files (x86)\Microsoft VS Code" (
copy \\server.yapsda.org\Software\VSCodeSetup-stable.exe C:\Temp\
VSCodeSetup-stable.exe /VERYSILENT /NORESTART
)

If Not Exist "C:\Program Files (x86)\Tipp10" (
copy \\server.yapsda.org\Software\tipp10_win_v2-1-0.exe C:\Temp\
tipp10_win_v2-1-0.exe /VERYSILENT /NORESTART
)

REM
REM Install SP1 if not installed
REM
wmic.exe os get servicepackmajorversion | findstr /B /C:"1" && (
If Not Exist "C:\Program Files (x86)\Microsoft Visual Studio 14.0" (
robocopy "\\server.yapsda.org\Software\Community 2015" C:\Temp\ /E
start /wait vs_community.exe /AdminFile "C:\Temp\AdminDeployment.xml" /Quiet /NoWeb /NoRestart
) ) || (
copy \\server.yapsda.org\Software\windows6.1-KB976932-X64.exe C:\Temp\
windows6.1-KB976932-X64.exe /unattend
)

REM
REM Remove temporary directory
REM
cd C:\
rmdir /S /Q C:\Temp
