
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
$TergetDirectory = "C:\DataCheck"
$PSFullName = Join-Path $TergetDirectory "DataChk.ps1"
if( -not (Test-Path $TergetDirectory)){mkdir $TergetDirectory}
Invoke-WebRequest -Uri https://raw.githubusercontent.com/MuraAtVwnet/ystsp/master/DataChk.ps1 -OutFile $PSFullName
$WsShell = New-Object -ComObject WScript.Shell
$Shortcut = $WsShell.CreateShortcut("C:\Users\Public\Desktop\ORS DataCheck.lnk")
$Shortcut.TargetPath = $TergetDirectory
$Shortcut.Save()
Invoke-Item $TergetDirectory
clsoe
