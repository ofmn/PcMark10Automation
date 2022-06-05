#region Install nessecary modules to pass data to SQL server
if (Get-Module -ListAvailable -Name Sqlserver, SQLPS) {
    Write-Host -ForegroundColor Green "SQL Module exists.. Continuing.."
    Write-Host
}
else {
    Write-Host -ForegroundColor Red "SQL Module does not exist"
    Write-Host -ForegroundColor Cyan "Installing..."
    Install-PackageProvider -Name NuGet -Scope CurrentUser -MinimumVersion 2.8.5.201 -Confirm:$False -Force
    Install-Module -Name "Sqlserver" -Scope CurrentUser -Confirm:$False -Force
    Write-Host -ForegroundColor Green "Complete."

}
#endregion

#region Check for PCMark software - install if not present
$software = "*PCMark 10*";
$installed = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like $software })

If (-Not $installed) {
    Write-Host "'$software' is NOT installed.
Will install now";
    $pcmarkpath = (Get-ChildItem \\it\Operations\GLOPAS\Install\Automation\InstallFiles -Directory).FullName
    Start-Process -FilePath $pcmarkpath\pcmark10-setup.exe -Wait -ArgumentList "/c, /force, /silent"

}
else {
    Write-Host "'$software' is installed."
}

#endregion

#region functions

# Function to ask the console what benchmark to run.
# Also setting $deffile and $defname here - which are used to identify what benchmark we are running, where to put the files, and how to collect the scores in other functions.
function whatpcmark10 {
    $hashpcmark10 = [Ordered]@{
        "1 - Normal"              = "pcm10_benchmark.pcmdef"
        "2 - OfficeApps"          = "pcm10_applications.pcmdef"
        "3 - BatteryLifeOffice"   = "pcm10_applications_batterylife.pcmdef"
        "4 - Storage Performance" = "pcm10_storage_full_default.pcmdef"
        "5 - Express"             = "pcm10_express.pcmdef"
    }
    Write-Host 
    foreach ($key in $hashpcmark10) {
        $key.Keys
    }

    # Asking what def file we want to choose
    do {
        $answer = Read-Host "Press the number for the benchmark you want to run"
    } until ($answer -match "^1$|^2$|^3$|^4$|^5$")

    [int]$answer1 = $answer - 1
    $script:deffile = $hashpcmark10.Get_Item($answer1)
    $script:defname = $hashpcmark10.GetEnumerator() | Where-Object Name -match $answer | ForEach-Object Key
    $script:defname = $defname -replace '\d\s-\s', ''

    do {
        $script:loop = read-host "Please input number of loops you want to run, 0 = until stopped manually."
    } until ($loop -match "^[0-9]$")

    Remove-Item -Path C:\TempMark -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path "C:\" -Name "TempMark\$model\PCMark10\$defname" -ItemType "directory" | Out-Null

}
# Function to make sure we are waiting for the benchmark files/XML's to exist before moving on with the script (Also matching the amount of loops we specified)
function waitforfiles {
    param($ext, $loop)
    do {
        start-sleep -Seconds 5
        $script:files = Get-ChildItem "C:\TempMark\$model\PCMark10\$defname\*.$ext"
    } While ($files.Count -ne $loop)   
}
function startpcmark10 {
    param($Server, $model, $deffile, $defname, [int]$loop, $gettime)
    & \\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\HWiNFO64.EXE "-l\\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\LogFiles\$Server-$gettime.csv"
    Start-Sleep -s 5
    Set-Location "C:\Program Files\UL\PCMark 10"

    .\PCMark10Cmd.exe --register=PCM10-TPRO-20220824-2ZQQ6-PJD67-RHKU9-DW6JP --loop $loop --systeminfo=off --log="C:\Temp\pcmark10-$gettime.log" --definition=$deffile --export-xml="C:\TempMark\$model\PCMark10\$defname\$Server-$gettime-$defname.pcmark10-result.xml"

    waitforfiles -ext "pcmark10-result.xml" -loop $loop
    Start-Sleep -Seconds 5
    Stop-Process -Name HWiNFO64

    #Calls function to retrieve scores just made in this run - and will also be run in the end to fetch scores already existing
    update_scores -deffile $deffile -model $model -beforeorafter ThisRun

    Copy-Item -Path "c:\TempMark\$model\PCMark10\$defname\*" -Destination "\\it\Operations\GLOPAS\Public\BenchmarkingScores\$model\PCMark10\$defname" -Recurse
}
function getinfo {
    $script:Server = hostname.exe 
    $script:time = Get-Date -Format "dd-MM-yy.HH;mm" -ErrorAction SilentlyContinue
    $script:model = (Get-Wmiobject -class Win32_ComputerSystemProduct).Version
    If ($model -eq "System Version") { $script:model = "ThinkPad T14 Gen 2i" }

}
# New and improved update scores that dynamically work with what deffile is chosen - auto loops through provided path for XML's to create hashtable with scores we want to collect and then pass to SQL
# This code alone saved more than 430 lines of very bad code from previous version
function update_scores {
    param($deffile, $model, $beforeorafter)
    If ($deffile -eq "pcm10_benchmark.pcmdef") { $script:whatbenchmark = "Normal" }
    If ($deffile -eq "pcm10_applications.pcmdef") { $whatbenchmark = "OfficeApps" }
    If ($deffile -eq "pcm10_applications_batterylife.pcmdef") { $whatbenchmark = "BatteryLifeOffice" }
    If ($deffile -eq "pcm10_storage_full_default.pcmdef") { $whatbenchmark = "Storage Performance" }
    If ($deffile -eq "pcm10_express.pcmdef") { $whatbenchmark = "Express" }

    # Get all files from \\it\Operations\GLOPAS\Public\BenchmarkingScores\$model
    If ($beforeorafter -eq "OtherRuns") { $xmlfiles = Get-ChildItem -Path "\\it\Operations\GLOPAS\Public\BenchmarkingScores\$model\PCMark10\$whatbenchmark" -Filter *.xml -Recurse }
    If ($beforeorafter -eq "ThisRun") { $xmlfiles = Get-ChildItem -Path "C:\TempMark\$model\PCMark10\$whatbenchmark" -Filter *.xml -Recurse }

    # If statement to skip this if count of XML's is 0 (Like for a brand new model)
    If ($xmlfiles.Count -ne 0) {
        [xml]$XmlDocument2 = Get-Content $xmlfiles[0].FullName

        # The XML files containing scores is in two parts, for all the different benchmarks.
        # Unfortunetaly depending on what benchmark is chosen we need to collect the scores from different parts (Either 1 or second)
        # Below two lines get the different score name from each and put it into an array
        # For "Normal, Express, OfficeApps" we have to use the second part when we collect scores, or else we do not retrieve the more specific scores like Word, Excel, VideoConf etc
        # For Battery, and Storage we need to use the first part.
        # This is why we point to either 0/1 in below lines

        # Load the first xml file based on what benchmark is chosen - and use a random xml to dynamically create the scores we want to collect.
        # Below 3 If statements is specifically only to collect the NAMES of the scores
        # Then we set what $partInXML we want to collect the actual scores for later down the script.
        If ($whatbenchmark -eq "BatteryLifeOffice") {
            $selectscore = "*PCMark10BatterylifeApplicationsRuntime*"
            $script:getscorenames0 = $XmlDocument2.benchmark.results.result[0] | Select-Object $selectscore
            $partInXML = 0   
        }

        If ($whatbenchmark -eq "Normal" -or $whatbenchmark -eq "OfficeApps" -or $whatbenchmark -eq "Express") {
            $selectscore = "*score" 
            $script:getscorenames0 = $XmlDocument2.benchmark.results.result[1] | Select-Object $selectscore
            $partInXML = 1 
        }

        If ($whatbenchmark -like "Storage Performance") {
            $selectscore = "*Score", "Pcm10StorageFullAverageAccessTimeOverall", "Pcm10StorageFullBandwidthOverall"  
            $script:getscorenames0 = $XmlDocument2.benchmark.results.result[0] | Select-Object $selectscore
            $partInXML = 0
        }

        # Create empty hashtable
        $script:hashtable = @{}

        # Foreach score found in the first array we just did, use its name example "writingscore" as the $key
        # And the value as the same as the key, but with $ infront, so we can add the values of all scores to variable, that has the same name as the columns in SQL, where we want to add them to eventually
        foreach ($item in $getscorenames0.psobject.properties) {
            $hashtable.Add($item.Name, "$" + $item.Name)
        }

        # Creating the actually array to put all the number scores into
        $global:scorearray0 = @()
        # Foreach to loop through all xml files found
        foreach ($entry in $xmlfiles) {
            # Set full path to each of the xml files
            $path1 = $entry.Fullname
            # Get content of XML 
            [xml]$script:XmlDocument3 = Get-Content $path1
            # Select-Object filter from above to only chose scores that matches the regex
            $script:scorearray0 += $XmlDocument3.benchmark.results.result[$partInXML] | Select-Object $selectscore
    
        }

        # Foreach $key in hashtable, collect all scores from the scorearray and average them, and round to 0 decimal - put into $hashtable $key, which are also variables
        # $key below meaning the different scores we are collecting, VideoConf, Writing, Spreadsheets etc
        foreach ($key in $($hashtable.Keys)) {
            If ($key -match ".*score|PCMark10.*|EssentialsScore|ProductivityScore|DigitalContentCreationScore|Pcm10StorageFull.*") {

                # Special case for BatteryLife as it needs to be converted from seconds to minutes (divide by 60)
                If ($key -eq "PCMark10BatteryLifeApplicationsRunTime") { $hashtable[$key] = [math]::Round(($scorearray0.PCMark10BatteryLifeApplicationsRunTime / 60 | Measure-Object -Average).Average) }
                else {
                    $hashtable[$key] = [math]::Round(($scorearray0.$key | Measure-Object -Average).Average)
                }
            }
        }
    }
    If ($beforeorafter -eq "ThisRun") { $script:thisrun = $hashtable }
    If ($beforeorafter -eq "OtherRuns") { $script:otherruns = $hashtable }
}

#endregion

getinfo

whatpcmark10

## Import Maximum performance powerplan and set it active
powercfg /import "\\it\Operations\GLOPAS\Install\Tools\Power Plan\highperformance.pow" 6215c520-3670-4349-89e2-186b0dca3999
powercfg /setactive 6215c520-3670-4349-89e2-186b0dca3999

## Start PCMark 10
startpcmark10 -Server $Server -model $model -deffile $deffile -defname $defname -loop $loop -gettime $time

## Get scores a second time, to compare current run VS all runs for same model
update_scores -deffile $deffile -model $model -beforeorafter OtherRuns


#Create Table object
$Result = New-Object system.Data.DataTable “TestTable”

#Define Columns
$BName = New-Object system.Data.DataColumn BenchMark, ([string])
$Thisruncolumn = New-Object system.Data.DataColumn ThisRun, ([string])
$OtherRunscolumn = New-Object system.Data.DataColumn AllRuns, ([string])
$Difference = New-Object system.Data.DataColumn Difference, ([string])

#Add the Columns
$Result.columns.add($BName)
$Result.columns.add($Thisruncolumn)
$Result.columns.add($OtherRunscolumn)
$Result.columns.add($Difference)

foreach ($item in $thisrun.Keys) {
    #Create a row
    $row = $Result.NewRow()
    $thisrunscore = $thisrun.Get_Item($item)
    $otherrunsscore = $otherruns.Get_Item($item)
    $differencescore = $thisrunscore - $otherrunsscore
    #Add to row
    $row.BenchMark = "$item"
    $row.ThisRun = $thisrunscore
    $row.OtherRuns = $otherrunsscore
    $row.Difference = $differencescore
    

    #Add row to table
    $Result.Rows.Add($row)
}

$Result

\\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\LogViewer\GenericLogViewer.exe "\\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\LogFiles\$Server-$time.csv"