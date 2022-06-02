Clear-Host
# URL that always point to latest Pcmark version
$Uri = 'http://benchmarks.ul.com/downloads/pcmark10-professional.zip'

[regex] $regex = "PCMark10.*professional.zip"

# Code to get the redirected link, which contains the version of the latest PCMark 10 build and match it against regex
$found = (Invoke-WebRequest -Method Get -Uri $Uri -MaximumRedirection 0 -ErrorAction SilentlyContinue).Links.href -match $regex
if ($found) { $pcmarkversion = $matches.Values }


# Set path for new PCMark zip/file if newer verison is found
$Filepath = "\\it\Operations\GLOPAS\Install\Automation\InstallFiles\$pcmarkversion"
If (!(Test-Path $Filepath)) {

$webClient = New-Object System.Net.WebClient
$Webclient.DownloadFile($Uri, "$Filepath")

$basename = $pcmarkversion -replace ".zip",""

Write-host -ForegroundColor Green "Done downloading new PCMark10 to $Filepath"

Expand-Archive $Filepath -DestinationPath \\it\Operations\GLOPAS\Install\Automation\InstallFiles\$basename
}


# Check if double folder exist, if two versions are found, delete oldest versions folder and its corresponding zip
$Checkfordouble = Get-ChildItem \\it\Operations\GLOPAS\Install\Automation\InstallFiles -Directory

If ($Checkfordouble.count -gt 1) {


$oldestversion = $Checkfordouble | Sort-Object lastwritetime | Select-Object -First 1
Write-host "More than two versions found
Deleting older version - $oldestversion"


$version2delete = $oldestversion.FullName
Remove-Item -Path $version2delete -Recurse -Force
Remove-Item -Path $version2delete".zip" -Recurse -Force
}