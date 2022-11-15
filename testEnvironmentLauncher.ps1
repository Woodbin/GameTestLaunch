<#
 Game Test Environment Launcher
 ===================================
 This script requires the following software:

 OBSStudio with OBS working directory set in $obsWd
 LibreOffice Writer/MSWord set in $librePath
 SysInternals RamMap set in $rmapPath
 SysInternals Process Explorer set in $procexpPath
 SysInternals VmMap set in $vmapPath
 ShareX set in $screenshotAppPath
 -----------------------------------
 A tool for quick launching your game testing session.
 Starts up OBS Studio, Writer, Notepad, RamMap, VmMap, Process Explorer.
 Also opens bugtracking page specified in $bugtracking
 Also opens timetracking pace specified in $timetracking
 Prints out info about memory, pagefile use, remaining drive free space. 
 Scripts checks if utilities are running before launching them, preventing multiple instances

 You can switch between using ShareX/Lightshot and notepad/notepad++ 
 by commenting lines in and out

 REWRITE DEFAULT PARAMETERS BETWEEN LAUNCHING!

#>

param (
# Refresh time of profiler
[int]$profilerRefreshRate = 10,
# Path to game executable
$gamePath = "path to game",
# Path to file with notes
$notes = "path to notes",
# URL of bugtracking service
$bugtracking = "http://url.to.bugtracking",
# URL of timetracking service
$timetracking = "http://timetracking.app",
# OBS working directory
$obsWd = "path to obs directory",
# Paths to utility executables
$librePath = "path to swriter",
$rmapPath = "path to rammap",
$vmapPath = "path to vmmap",
$procexpPath = "path to procexp",
$notepadPath = "path to notepad++",
$notepadProcess = "notepad++",
#$notepadPath = notepad,
#$notepadProcess = notepad,
$screenshotAppPath = "path to shareX-launcher",
$screenshotAppProcessName = "ShareX",
#$screenshotAppPath = "path to lightshot",
#$screenshotAppProcessName = "lightshot",
# Launch the game during script start
$gamelaunch = $false,
# Open bugtracking
$launchBugtracking = $true,
# Open timetracking
$launchTimetracking = $true,
# Launch OBS
$obsLaunch = $true,
# Launch Writer
$writerLaunch = $true,
# Launch Notepad
$notepadLaunch = $true,
# Launch ScreenshotApp
$screenshotLaunch = $true,
# Launch RamMap
$rmapLaunch = $true,
# Launch VmMap 
$vmapLaunch = $true,
# Launch ProcExp
$procexpLaunch = $true,
# Script debug mode
$debug = $false
)

$startingValues = [ordered]@{
    time = "";
    sysDrive = "";
    gameDrive = "";
    sysPageFile = "";
    gamePageFile = "";
    sysMemory = "";
    cpuUse = "";
}

$values = [ordered]@{
    time = "";
    sysDrive = "";
    gameDrive = "";
    sysPageFile = "";
    gamePageFile = "";
    sysMemory = "";
    gameMemory = "";
    cpuUse = "";
}

$lastValues = [ordered]@{
    time = "";
    sysDrive = "";
    gameDrive = "";
    sysPageFile = "";
    gamePageFile = "";
    sysMemory = "";
    gameMemory = "";
    cpuUse = "";
}


function RefreshTimestamp{ $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss" }

function DrivesFreeSpace {
    $global:cdrive = (Get-PSDrive C)
    $global:gamedrive = (Get-PSDrive ($gameDisk))
    $sysfree = ($cdrive.Free)/1GB
    $gamefree = ($gamedrive.Free)/1GB
    Write-Host "Drive Free Space:`n=============================`n C:       $sysfree GB`n K:       $gamefree GB`n"
}
function PageFilesSize {
    $pages=Get-CimInstance Win32_PageFile | Select-Object Name, InitialSize, MaximumSize, Filesize
    Write-Host "Pagefiles:"
    Write-Host ( $pages | Format-Table | Out-String)
}
function refreshPids {
    $obsActive = Get-Process obs64 -ErrorAction SilentlyContinue
    if($null -eq $obsActive){
    if($obsLaunch){
    #process null when should be running
    }
    }else{
    $obs = Get-Process obs64
    }
    $global:obsId = ($obs.Id)

    $writerActive = Get-Process $libreProcess -ErrorAction SilentlyContinue
    if($null -eq $writerActive){
    if($writerLaunch){
    #process null when should be running
    }
    }
    else{
    $writer = Get-Process $libreProcess
    }
    $global:writerId = ($writer.Id)

    $rmapActive = Get-Process RAMMap64 -ErrorAction SilentlyContinue
    if($null -eq $rmapActive) {
    #process null when should be running
    }
    else{
    $rmap = Get-Process RAMMap64
    }
    $global:rmapId = ($rmap.Id)

    $procexpActive = Get-Process procexp64 -ErrorAction SilentlyContinue
    if($null -eq $procexpActive){ 
    if($procexpLaunch){
    #process null when should be running
    }
    }
    else{
    $procexp = Get-Process procexp64
    }
    $global:procexpId = ($procexp.Id)

    $notepadActive = Get-Process $notepadProcess -ErrorAction SilentlyContinue
    If($null -eq $notepadActive){
        if($notepadLaunch){
        #process null when should be running
        }
    }else{
    $notepad = Get-Process $notepadProcess
    }
    $global:notepadId = ($notepad).Id

    $screenshotActive = Get-Process $screenshotAppProcessName -ErrorAction SilentlyContinue
    if($null -eq $screenshotActive) {
        if($screenshotLaunch){
        #process null when should be running
        }
    }
    else{
    $screenshot = Get-Process $screenshotAppProcessName
    }
    $global:screenshotId = ($screenshot.Id)

    $vmapActive = Get-Process vmmap64 -ErrorAction SilentlyContinue
    if($null -eq $vmapActive) { 
        if($vmapLaunch){
        #process null when should be running
        }
        }else{
        $vmap = Get-Process vmmap64
    }

    $global:vmapId = ($vmap.Id)

    if($debug){
    Write-Host "refreshPids called"
    Write-Host "$gameId,$obsId,$writerId,$notepadId,$screenshotId,$rmapId,$vmapId"
}

}
function CreatePidTable {
    #Clear-Variable PIDtable
    $global:PIDtable = New-Object System.Data.DataTable
       refreshPids
    #[void]$PIDtable.Clear()
    [void]$global:PIDtable.Columns.Add("$gameName")
    [void]$global:PIDtable.Columns.Add("OBS")
    [void]$global:PIDtable.Columns.Add("Writer")
    [void]$global:PIDtable.Columns.Add("Notepad")
    [void]$global:PIDtable.Columns.Add("Screenshot app")
    [void]$global:PIDtable.Columns.Add("RamMap")
    [void]$global:PIDtable.Columns.Add("VmMap")
    #[void]$PIDtable.Rows.Add("$gameName", "OBS","Writer","Notepad","Screenshot app", "RamMap", "VmMap")
    [void]$global:PIDtable.Rows.Add($gameId,$obsId,$writerId,$notepadId,$screenshotId,$rmapId,$vmapId)
    if($debug){
    Write-Host "CreatePidTable called"
    Write-Host ($global:PIDtable | Format-Table | Out-String)
    }
}

function PIDs {
    CreatePidTable
    Write-Host "Process IDs:"
    Write-Host ($global:PIDtable | Format-Table | Out-String)
}

function getStartValues(){
    $cdrive = (Get-PSDrive C)
    $gamedrive = (Get-PSDrive ($gameDisk))
    $pages=Get-CimInstance Win32_PageFile | Select-Object Name, InitialSize, MaximumSize, Filesize
    $sysfree = (($cdrive.Free)/1GB).ToString("N")
    $sysmax = ((($cdrive.Free)+($cdrive.Used))/1GB).ToString("N")
    $gamefree = (($gamedrive.Free)/1GB).ToString("N")
    $gamemax = ((($gamedrive.Free)+($gamedrive.Used))/1GB).ToString("N")
    $sysMem = $os | Select-Object @{Name = "FreeGB";Expression = {[math]::Round($_.FreePhysicalMemory/1mb,2)}},    
    @{Name = "TotalGB";Expression = {[int]($_.TotalVisibleMemorySize/1mb)}}
    $cpuLoad = (Get-CimInstance -ClassName win32_processor | Measure-Object -Property LoadPercentage -Average).Average
    $startingValues["time"] = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $startingValues["sysDrive"] = "Free: $sysfree GB / Max:$sysmax GB"
    $startingValues["gameDrive"] = "Free: $gamefree GB / Max: $gamemax GB"
    $startingValues["sysPageFile"] =  "Initial: $($pages.get(0).InitialSize) MB / Current: $(($pages.Get(0).FileSize)/1MB) MB / Max: $($pages.Get(0).MaximumSize) MB"
    $startingValues["gamePageFile"] = "Initial: $($pages.get(1).InitialSize) MB / Current: $(($pages.Get(1).FileSize)/1MB) MB / Max: $($pages.Get(1).MaximumSize) MB"
    $startingValues["sysMemory"] = "Free: $($sysMem.FreeGB) GB / Total: $($sysMem.TotalGB) GB"
    $startingValues["cpuUse"] = "CPU Load: $cpuLoad % | Cores: $cores Logical Processors: $logical Max Clock: $clock MHz"
    if($debug){
        Write-Host "Starting Values:"
        Write-Host ($startingValues | Format-Table | Out-String)
    }
}

function getValues(){
    $cdrive = (Get-PSDrive C)
    $gamedrive = (Get-PSDrive ($gameDisk))
    $pages=Get-CimInstance Win32_PageFile | Select-Object Name, InitialSize, MaximumSize, Filesize
    $sysfree = (($cdrive.Free)/1GB).ToString("N")
    $sysmax = ((($cdrive.Free)+($cdrive.Used))/1GB).ToString("N")
    $gamefree = (($gamedrive.Free)/1GB).ToString("N")
    $gamemax = ((($gamedrive.Free)+($gamedrive.Used))/1GB).ToString("N")
    $sysMem = $os | Select-Object @{Name = "FreeGB";Expression = {[math]::Round($_.FreePhysicalMemory/1mb,2)}},
    @{Name = "TotalGB";Expression = {[int]($_.TotalVisibleMemorySize/1mb)}}
    $cpuLoad = (Get-CimInstance -ClassName win32_processor | Measure-Object -Property LoadPercentage -Average).Average
    $game = Get-Process $gameName -ErrorAction SilentlyContinue
    $gameCpuLoad = (Get-WmiObject -class Win32_PerfFormattedData_PerfProc_Process | Where-Object {$_.Name -eq "$gameName"}).PercentProcessorTime
    if(!($game)){
    $gameWs = 0
    $gamePm = 0    
    }Else{
    $gameWs = (((Get-Process $gameName).WorkingSet64)/1MB)
    $gamePm = (((Get-Process $gameName).PrivateMemorySize64)/1MB)
    }
    $gamemem = "Working Set: $gameWs MB | Private Memory: $gamePm MB"
    $values["time"] = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $values["sysDrive"] = "Free: $sysfree GB / Max:$sysmax GB"
    $values["gameDrive"] = "Free: $gamefree GB / Max: $gamemax GB"
    $values["sysPageFile"] =  "Initial: $($pages.get(0).InitialSize) MB / Current: $(($pages.Get(0).FileSize)/1MB) MB / Max: $($pages.Get(0).MaximumSize) MB"
    $values["gamePageFile"] = "Initial: $($pages.get(1).InitialSize) MB / Current: $(($pages.Get(1).FileSize)/1MB) MB / Max: $($pages.Get(1).MaximumSize) MB"
    $values["sysMemory"] = "Free: $($sysMem.FreeGB) GB / Total: $($sysMem.TotalGB) GB"
    $values["cpuUse"] = "CPU Load: $cpuLoad % $gameName Usage: $gameCpuLoad % | Cores: $cores Logical Processors: $logical Max Clock: $clock MHz"
    $values["gameMemory"] = $gamemem
    if($debug){
        Write-Host "Values:"
        Write-Host ($values | Format-Table | Out-String)
    }
}

function exportReport(){
    $timeStamp = Get-Date -Format "MM/dd/yyyy HH-mm-ss"
    $filename = "$gameName $timeStamp.txt"
    New-Item -ItemType File -Path $wd -Name $filename
    Add-Content -Path "$wd\$filename" -Value "$gameName profiling report" 
    Add-Content -Path "$wd\$filename" -Value "Starting values:"
    Add-Content -Path "$wd\$filename" -Value ($startingValues | Format-Table | Out-String)
    Add-Content -Path "$wd\$filename" -Value "Last recorded:"
    Add-Content -Path "$wd\$filename" -Value ($lastValues | Format-Table | Out-String)
    Add-Content -Path "$wd\$filename" -Value "After exit:"
    Add-Content -Path "$wd\$filename" -Value ($values | Format-Table | Out-String)
}

function backupValues(){
    $lastValues["time"] = $values["time"]
    $lastValues["sysdrive"] = $values["sysdrive"]
    $lastValues["gameDrive"] = $values["gameDrive"]
    $lastValues["sysPageFile"] = $values["sysPageFile"]
    $lastValues["gamePageFile"] = $values["gamePageFile"]
    $lastValues["sysMemory"] = $values["sysMemory"]
    $lastValues["gameMemory"] = $values["gameMemory"]
    $lastValues["cpuUse"] = $values["cpuUsed"]
}

function refreshDrives(){
    $cdrive = (Get-PSDrive C)
    $gamedrive = (Get-PSDrive ($gameDisk))
}
function ProfileGame {
    $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    Write-Host "Resources at $timeStamp"
    PIDs
    backupValues
    getValues
    Write-Host ($values | Format-Table | Out-String)
    #Write-Host "Memory Info:"
    #Write-Host ($memory| Format-Table | Out-String)
    #Write-Host "Game memory usage: $gameMemory" 
    #DrivesFreeSpace
    #PageFilesSize
}

if($debug){
    Write-Host "Script starting. Work directory:"
}
$wd = Get-Location

if($debug){
    Write-Host $wd
}


$timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
# Name of game process
$gameName = (Get-ChildItem $gamePath).BaseName
$libreProcess = (Get-ChildItem $librePath).BaseName
$gameDisk = $gamePath[0]
$gameLaunchWaitTime = 10
$PIDtable = New-Object System.Data.DataTable
Write-Host "Script Started at $timestamp"
$cdrive = (Get-PSDrive C)
$gamedrive = (Get-PSDrive ($gameDisk))
$os = Get-Ciminstance Win32_OperatingSystem
$pages=Get-CimInstance Win32_PageFile | Select-Object Name, InitialSize, MaximumSize, Filesize
$cores = (Get-CimInstance Win32_Processor).NumberOfCores
$logical = (Get-CimInstance Win32_Processor).NumberOfLogicalProcessors
$clock = (Get-CimInstance Win32_Processor | Select-Object -ExpandProperty MaxClockSpeed)
# Open bugtracking
if($launchBugtracking){ Start-Process $bugtracking}
# Open timetracking
if($launchTimetracking) { Start-Process $timetracking }
# Launch OBS

$obsActive = Get-Process obs64 -ErrorAction SilentlyContinue
if($null -eq $obsActive){
 if($obsLaunch){$obs = Start-Process "$obsWd\obs64.exe" -PassThru -WorkingDirectory $obsWd}
 }else{
 $obs = Get-Process obs64
 }
$obsId = ($obs.Id)
CreatePidTable
If ($debug) { Write-Host "OBS PID: $obsId"}

# Launch LibreOffice Writer
$writerActive = Get-Process $libreProcess -ErrorAction SilentlyContinue
if($null -eq $writerActive){
 if($writerLaunch){$writer = Start-Process $librePath -PassThru}
 }
else{
 $writer = Get-Process $libreProcess
 }
$writerId = ($writer.Id)
CreatePidTable
If ($debug) { Write-Host "Writer PID: $writerId" }

#Launch RamMap
$rmapActive = Get-Process RAMMap64 -ErrorAction SilentlyContinue
if($null -eq $rmapActive) {
 if($rmapLaunch){$rmap = Start-Process $rmapPath -PassThru}
 }
else{
 $rmap = Get-Process RAMMap64
 }
$rmapId = ($rmap.Id)
CreatePidTable
If ($debug) { Write-Host "RamMap PID: $rmapId" }

# Launch Process Explorer
$procexpActive = Get-Process procexp64 -ErrorAction SilentlyContinue
if($null -eq $procexpActive){ 
if($procexpLaunch){$procexp = Start-Process $procexpPath -PassThru}
}
else{
 $procexp = Get-Process procexp64
 }
$procexpId = ($procexp.Id)
CreatePidTable
If ($debug) { Write-Host "Procexp PID: $procexpId" }

# Launch notepad for notes
$notepadActive = Get-Process $notepadProcess -ErrorAction SilentlyContinue
If($null -eq $notepadActive){
If(Test-Path -Path $notes -PathType Leaf){
    if($notepadLaunch){
    $notepad = Start-Process $notepadPath -PassThru -ArgumentList "$notes"
    }
}else{
    if($notepadLaunch){
    $notepad = Start-Process $notepadPath -PassThru}
    }
}
else{
 $notepad = Get-Process $notepadProcess
 }
$notepadId = ($notepad.Id)
CreatePidTable
If ($debug) { Write-Host "Notepad PID: $notepadId"}

# Launch Screenshot app
$screenshotActive = Get-Process $screenshotAppProcessName -ErrorAction SilentlyContinue
if($null -eq $screenshotActive) {
    if($screenshotLaunch){
 $screenshot = Start-Process $screenshotAppPath -PassThru}
 }
else{
 $screenshot = Get-Process $screenshotAppProcessName
 }
$screenshotId = ($screenshot.Id)
CreatePidTable
If ($debug) { Write-Host "Screenshot app PID: $screenshotId"}

# Return to script working directory
Set-Location -Path $wd

# Print resources at start
#DrivesFreeSpace
#PageFilesSize
getStartValues
Write-Host "Starting values:"
Write-Host ($startingValues | Format-Table | Out-String)
CreatePidTable
PIDs
#Try to launch the game
if($gameLaunch){Start-Process $gamePath}

# Wait for game to start
$started = $false
Write-Host 'Waiting for game to start'
Do {
$game = Get-Process $gameName -ErrorAction SilentlyContinue
   If (!($game)) { Start-Sleep -Seconds $gameLaunchWaitTime }
   Else{ 
   $gameId = (Get-Process $gameName).Id
   Write-Host "$gameName running, pid: $gameId"
   $started = $true 
   }
}
Until($started)
$game = (Get-Process $gameName).Id
CreatePidTable
If($debug) { Write-Host "Game PID: $game"}

#Launch VMMap
$vmapActive = Get-Process vmmap64 -ErrorAction SilentlyContinue
if($vmapActive) { $vmapActive | Stop-Process -Force }
if($vmapLaunch){$vmap = Start-Process "$vmapPath" -PassThru -ArgumentList "-p $game"}
$vmapId = ($vmap.Id)
CreatePidTable
If ($debug) { Write-Host "VmmMap PID: $vmapId" }

# Profiling Loop
$gameRunning = (Get-Process $gameName -ErrorAction SilentlyContinue)
Do {
    Clear-Host
    ProfileGame
    Start-Sleep -Seconds $profilerRefreshRate
    $gameRunning = (Get-Process $gameName -ErrorAction SilentlyContinue)
}
While($gameRunning)

#Resources after game stops running
$timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
Write-Host "[$timestamp] Game stopped running!"
#DrivesFreeSpace
#PageFilesSize
Write-Host "Starting values:"
Write-Host ($startingValues | Format-Table | Out-String)
Write-Host "Last Recorded while running: "
Write-Host ($lastValues | Format-Table | Out-String)
getValues
Write-Host "Values at finish:"
Write-Host ($values | Format-Table | Out-String)
CreatePidTable
PIDs
exportReport
