#region Install nessecary modules to pass data to SQL server
if (Get-Module -ListAvailable -Name Sqlserver,SQLPS) {
    Write-Host -ForegroundColor Green "SQL Module exists.. Continuing.."
    Write-Host
} else {
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

If(-Not $installed) {
	Write-Host "'$software' is NOT installed.
Will install now";
$pcmarkpath = (Get-ChildItem \\it\Operations\GLOPAS\Install\Automation\InstallFiles -Directory).FullName
Start-Process -FilePath $pcmarkpath\pcmark10-setup.exe -Wait -ArgumentList "/c, /force, /silent"

} else {
	Write-Host "'$software' is installed."}

#endregion

#region functions
function whatpcmark10 {
    $hashpcmark10 = [Ordered]@{
    "1 - Normal" = "pcm10_benchmark.pcmdef"
    "2 - OfficeApps" = "pcm10_applications.pcmdef"
    "3 - BatteryLifeOffice" = "pcm10_applications_batterylife.pcmdef"
    "4 - Storage Performance" = "pcm10_storage_full_default.pcmdef"
    "5 - Express" = "pcm10_express.pcmdef"
    }
    Write-Host 
    foreach ($key in $hashpcmark10) {
    $key.Keys
    }

    # Asking what def file we want to choose
    do {
    $answer = Read-Host "Press the number for the benchmark you want to run"
    } until ($answer -match "^1$|^2$|^3$|^4$|^5$")

    [int]$answer1 = $answer -1
    $script:deffile = $hashpcmark10.Get_Item($answer1)
    $script:defname = $hashpcmark10.GetEnumerator() | Where-Object Name -match $answer | ForEach-Object Key
    $script:defname = $defname -replace '\d\s-\s',''

    do
    {
    $script:loop = read-host "Please input number of loops you want to run, 0 = until stopped manually."
    } until ($loop -match "^[0-9]$")

    Remove-Item -Path C:\TempMark -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path "C:\" -Name "TempMark\$model\PCMark10\$defname" -ItemType "directory" | Out-Null

}
function waitforfiles {
    param($ext, $loop)
    do {
        start-sleep -Seconds 5
        $script:files = Get-ChildItem "C:\TempMark\$model\PCMark10\$defname\*.$ext"
        } While ($files.Count -ne $loop)   
}
function startpcmark10 {
param($Server, $model,$deffile,$defname,[int]$loop, $gettime)
& \\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\HWiNFO64.EXE "-l\\it\Operations\GLOPAS\Install\Scripts\SoftwareServices\HWINFO64\LogFiles\$Server-$gettime.csv"
Start-Sleep -s 5
Set-Location "C:\Program Files\UL\PCMark 10"
.\PCMark10Cmd.exe --register=PCM10-TPRO-20220824-2ZQQ6-PJD67-RHKU9-DW6JP --loop $loop --systeminfo=off --log="C:\Temp\pcmark10-$gettime.log" --definition=$deffile --out="C:\TempMark\$model\PCMark10\$defname\$Server-$gettime-$defname.pcmark10-result"

waitforfiles -ext "pcmark10-result" -loop $loop
Start-Sleep -Seconds 5
Stop-Process -Name HWiNFO64

foreach ($item in $files) {
& 'C:\Program Files\UL\PCMark 10\PCMark10Cmd.exe' --in="$item" --export-xml "$item.xml"
}

waitforfiles -ext "xml" -loop $loop

Copy-Item -Path "c:\TempMark\$model\PCMark10\$defname\*" -Destination "\\it\Operations\GLOPAS\Public\BenchmarkingScores\$model\PCMark10\$defname" -Recurse
}
function getinfo {
$script:Server = hostname.exe 
$script:time = Get-Date -Format "dd-MM-yy.HH;mm" -ErrorAction SilentlyContinue
$script:model = (Get-Wmiobject -class Win32_ComputerSystemProduct).Version
If ($model -eq "System Version") {$script:model = "ThinkPad T14 Gen 2i"}

}
# New and improved update scores that dynamically work with what deffile is chosen - auto loops through provided path for XML's to create hashtable with scores we want to collect and then pass to SQL
# This code alone saved more than 430 lines of very bad code from previous version
# In this function; implement what $deffile to use - and change the path to where this can be found - also based on the model
# $deffile = express | $model = P53 | Look for xml files matching express in P53 public folder
function update_scores {
    param($deffile, $model, $beforeorafter)
If ($deffile -eq "pcm10_benchmark.pcmdef") {$script:whatbenchmark = "Normal"}
If ($deffile -eq "pcm10_applications.pcmdef") {$whatbenchmark = "OfficeApps"}
If ($deffile -eq "pcm10_applications_batterylife.pcmdef") {$whatbenchmark = "BatteryLifeOffice"}
If ($deffile -eq "pcm10_storage_full_default.pcmdef") {$whatbenchmark = "Storage Performance"}

# Get all files from \\it\Operations\GLOPAS\Public\BenchmarkingScores\$model
# Probably make some logic in regards to what $deffile to use 
If ($beforeorafter -eq "OtherRuns") {$xmlfiles = Get-ChildItem -Path "\\it\Operations\GLOPAS\Public\BenchmarkingScores\$model\PCMark10\$whatbenchmark" -Filter *.xml -Recurse}
If ($beforeorafter -eq "ThisRun") {$xmlfiles = Get-ChildItem -Path "C:\TempMark\$model\PCMark10\$whatbenchmark" -Filter *.xml -Recurse}

# Load the first xml file in
[xml]$XmlDocument2 = Get-Content $xmlfiles[0].FullName

# The XML files is diveded into two different "result" 
# Below two lines get the different score name from each and put it into an array
# The first contains the overall score names, like "Pcmark10 overall", "Essentials", "Productivity"¨
# The second contains all the score names that are either under essentials, productivity or digital content
# Like, Browsing, VideoConf, Writing, App startup etc etc
If ($whatbenchmark -eq "BatteryLifeOffice") {
    $selectscore = "*PCMark10BatterylifeApplicationsRuntime*"
    $script:getscorenames0 = $XmlDocument2.benchmark.results.result[0] | Select-Object $selectscore   
}

If ($whatbenchmark -eq "Normal" -or $whatbenchmark -eq "OfficeApps") {
    $selectscore = "*score" 
    $script:getscorenames0 = $XmlDocument2.benchmark.results.result[0] | Select-Object $selectscore 
}

If ($whatbenchmark -like "Storage Performance") {
    $selectscore = "*Score", "Pcm10StorageFullAverageAccessTimeOverall", "Pcm10StorageFullBandwidthOverall"  
    $script:getscorenames0 = $XmlDocument2.benchmark.results.result[0] | Select-Object $selectscore
    
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
foreach ($entry in $xmlfiles) {
    $path1 = $entry.Fullname
    
    [xml]$script:XmlDocument3 = Get-Content $path1

    $script:scorearray0 += $XmlDocument3.benchmark.results.result[0] | Select-Object $selectscore
    
}

# Foreach $key in hashtable, collect all scores from the scorearray and average them, and round to 0 decimal - put into $hashtable $key, which are also variables
foreach ($key in $($hashtable.Keys)) {
    If ($key -match "PCMark10.*|EssentialsScore|ProductivityScore|DigitalContentCreationScore|Pcm10StorageFull.*") {

        # Special case for BatteryLife as it needs to be converted from seconds to minutes (divide by 60)
        If ($key -eq "PCMark10BatteryLifeApplicationsRunTime") { $hashtable[$key] = [math]::Round(($scorearray0.PCMark10BatteryLifeApplicationsRunTime / 60 | Measure-Object -Average).Average) }
        else {
            $hashtable[$key] = [math]::Round(($scorearray0.$key | Measure-Object -Average).Average)
        }
    }
}
If ($beforeorafter -eq "ThisRun") {$script:thisrun = $hashtable}
If ($beforeorafter -eq "OtherRuns") {$script:otherruns = $hashtable}
}

#endregion

getinfo

whatpcmark10

startpcmark10 -Server $Server -model $model -deffile $deffile -defname $defname -loop $loop -gettime $time

update_scores -deffile $deffile -model $model -beforeorafter ThisRun

update_scores -deffile $deffile -model $model -beforeorafter OtherRuns


#Create Table object
$Result = New-Object system.Data.DataTable “TestTable”

#Define Columns
$BName = New-Object system.Data.DataColumn BenchMark,([string])
$Thisruncolumn = New-Object system.Data.DataColumn ThisRun,([string])
$OtherRunscolumn = New-Object system.Data.DataColumn OtherRuns,([string])
$Difference = New-Object system.Data.DataColumn Difference,([string])

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