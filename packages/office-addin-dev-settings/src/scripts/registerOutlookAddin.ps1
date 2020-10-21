if ($args.Count -ne 1) {
    throw "Usage: sideloadOutlook.ps1 <manifestFilePath>"
}

$addinManifest = $args[0]
Write-Host 'Registering add-in manifest ' $addinManifest
New-App -FileData ([Byte[]](Get-Content -Encoding Byte -Path $addinManifest -ReadCount 0))