# Pre Setting
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force

# Get Script
$TergetDirectory = "C:\DataCheck"
$PSFullName = Join-Path $TergetDirectory "DataChk.ps1"
if( -not (Test-Path $TergetDirectory)){mkdir $TergetDirectory}
Invoke-WebRequest -Uri https://raw.githubusercontent.com/MuraAtVwnet/ystsp/master/DataChk.ps1 -OutFile $PSFullName

# Create shortcut
$WsShell = New-Object -ComObject WScript.Shell
$Shortcut = $WsShell.CreateShortcut("C:\Users\Public\Desktop\ORS DataCheck.lnk")
$Shortcut.TargetPath = $TergetDirectory
$Shortcut.Save()

# Open directory
Invoke-Item $TergetDirectory

