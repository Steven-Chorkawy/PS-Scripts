Clear-Host
Echo "Toggling ScrollLock..."  
$WShell = New-Object -com "Wscript.Shell" 
while ($true) { 
$WShell.sendkeys("{SCROLLLOCK}") 
# 30 seconds.
Start-Sleep -Milliseconds 30000   
$WShell.sendkeys("{SCROLLLOCK}") 
Start-Sleep -Seconds 100 
}