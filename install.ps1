## Konsole leeren
cls

## Cobra beenden
Taskkill /F /IM Cobra*

## Altes Plugin via WMIC entfernen
$MyApp = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -Like "cobra Outlook*"}
$MyApp.Uninstall()

## neues Plugin hinzufügen
msiexec.exe /i NMS64_Setup.msi /quiet

## Meldung
Write-Host "Installation abgeschlossen"
sleep 2

## Anwendung verlassen
Exit 0