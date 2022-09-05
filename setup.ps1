#Requires -RunAsAdministrator
$dl = "C:$Env:HOMEPATH\Downloads"
$sfdc_uri = "https://developer.salesforce.com/media/salesforce-cli/sfdx/channels/stable/sfdx-x64.exe"
$ps_uri = "https://github.com/PowerShell/PowerShell/releases/download/v7.2.6/PowerShell-7.2.6-win-x64.msi"

if ($(Write-Host "Download and install Powershell 7? (y/n): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) -eq "y"){
    Write-Host "Downloading Powershell 7..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$ps_uri" -OutFile "$dl\ps7.msi"
    Start-Process "$dl\ps7.msi" -Wait
}

if ($(Write-Host "Download SFDC CLI Package? (y/n): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) -eq "y"){
    Write-Host "Downloading SFDC CLI Package (sfdx)..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$sfdc_uri" -OutFile "$dl\sfdx-x64.exe" 
}

Write-Host "Launching SFDC CLI Setup (sfdx)... PLEASE WAIT - this process can take some time to launch and to complete" -ForegroundColor Yellow -BackgroundColor DarkGreen
Start-Process "$dl\sfdx-x64.exe" -Wait

Write-Host "Setting Powershell Execution Policy..." -ForegroundColor Yellow
Set-ExecutionPolicy Unrestricted

Write-Host "Done..." -ForegroundColor Green