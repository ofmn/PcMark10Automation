#Test commit from Ramboll
#Test commit from Home
#Test commit from Ramboll2
#test commit from home, with new "am i behind" enabled
Clear-Host
# URL that always point to latest Pcmark version
$Uri = 'http://benchmarks.ul.com/downloads/pcmark10-professional.zip'

[regex] $regex = "PCMark10.*professional.zip"

# Code to get the redirected link, which contains the version of the latest PCMark 10 build and match it against regex
$found = (Invoke-WebRequest -Method Get -Uri $Uri -MaximumRedirection 0 -ErrorAction SilentlyContinue).Links.href -match $regex
if ($found) { $pcmarkversion = $matches.Values }

# Set path for new PCMark zip/file if newer verison is found
$Filepath = "\\it\Operations\GLOPAS\Install\Automation\InstallFiles\$pcmarkversion"
#$Filepath = "\\it\Operations\GLOPAS\Install\Automation\InstallFiles\PCMark10-v2-1-2557-pro.zip"

If (!(Test-Path $Filepath)) {
    Write-Host "New version of PCMark10 found"
    Write-host "Press Y to make this script automatically download version $pcmarkversion, and use the new version instead."
    Write-Host "Press N to skip for this run, and use the current downloaded version."

    # Keep asking until we get a Yes or No.
    $answer = Read-Host "Yes, download or No, skip"

    while ("y.*|Y.*|n.*|N.*" -notmatch $answer) {
        $answer = Read-Host "Yes, download or No, skip"
    }
}


# Only proceed if there is a new version available, and user accepeted to download
If (!(Test-Path $Filepath) -and ($answer -match "y.*|Y.*")) {

    # Start downloading the file if it is found.
    $webClient = New-Object System.Net.WebClient
    $Webclient.DownloadFile($Uri, "$Filepath")

    $basename = $pcmarkversion -replace ".zip", ""

    Write-host -ForegroundColor Green "Done downloading new PCMark10 to $Filepath"

    # Unpack zip file with the same name, minus ".zip"
    Expand-Archive $Filepath -DestinationPath \\it\Operations\GLOPAS\Install\Automation\InstallFiles\$basename

    # Check if double folder exist, if two versions are found, delete oldest versions folder and its corresponding zip
    $Checkfordouble = Get-ChildItem \\it\Operations\GLOPAS\Install\Automation\InstallFiles -Directory

    # If more than one folder is found - it means we got at least two versions
    If ($Checkfordouble.count -gt 1) {
        # Get the lastwritetime from both, and select the oldest
        $oldestversion = $Checkfordouble | Sort-Object lastwritetime | Select-Object -First 1
        Write-host "More than two versions found
Deleting older version - $oldestversion"

        # Get the fullname of the oldest folder, delete it - and also delete corresponding zip.
        $version2delete = $oldestversion.FullName
        Remove-Item -Path $version2delete -Recurse -Force
        Remove-Item -Path $version2delete".zip" -Recurse -Force
    }

}
