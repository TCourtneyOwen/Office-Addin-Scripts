if ($args.Count -ne 1) {
    throw "Usage: sideloadOutlook.ps1 <manifestFilePath>"
}

$addinManifest = $args[0]
$modules = Get-InstalledModule
$moduleFound = $false;

for ($count = 0; $count -le $modules.Count; $count++)
{    
    if ($modules[$count].Name -eq "ExchangeOnlineManagement") {
        Write-Host 'Module found ' $modules[$count].Name
        $moduleFound = $true
        break
    } 
}

if ($moduleFound -ne $true) {
   Write-Host 'Installing ExchangeOnlineManagement module'
   Install-Module -Name ExchangeOnlineManagement
}

$sessions = Get-PSSession
if ($sessions.Count -ge 1)
{
    for ($count = 0; $count -le $sessions.Count; $count++)
    {
        if ($sessions[$count].Name -like "ExchangeOnlineInternalSession*") {
            break
        } else {
            Write-Host 'Connecting to ExchangeOnline'
            Connect-ExchangeOnline -ShowProgress $true
        }
    }
} else {
    Write-Host 'Connecting to ExchangeOnline'
    Connect-ExchangeOnline -ShowProgress $true
}

Write-Host 'Registering add-in manifest ' $addinManifest
New-App -FileData ([Byte[]](Get-Content -Encoding Byte -Path $addinManifest -ReadCount 0))