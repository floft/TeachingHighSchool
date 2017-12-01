REM Create System Restore Point
wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint "20160921Lab", 100, 7