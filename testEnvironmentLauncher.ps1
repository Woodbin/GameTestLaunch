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
 Prints out info about memory, pagefile use, remaining drive free space. 
 Scripts checks if utilities are running before launching them, preventing multiple instances

 You can switch between using ShareX/Lightshot and notepad/notepad++ 
 by commenting lines in and out

 REWRITE DEFAULT PARAMETERS BETWEEN LAUNCHING!

#>

param (
# Refresh time of profiler
[int]$profilerRefreshRate = 10,
# Launch the game during script start
$gamelaunch = $false,
# Path to game executable
$gamePath = "path to game",
# Path to file with notes
$notes = "path to notes",
# URL of bugtracking service
$bugtracking = "http:\\url.to.bugtracking",
$openBugtracking = $true,
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

# Script debug mode
$debug = $true
)

function RefreshTimestamp() { $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss" }

function DrivesFreeSpace {
    $cdrive = (Get-PSDrive C)
    $kdrive = (Get-PSDrive ($gameDisk))
    $cfree = ($cdrive.Free)/1GB
    $kfree = ($kdrive.Free)/1GB
    Write-Host "Drive Free Space:`n=============================`n C:       $cfree GB`n K:       $kfree GB`n"
}
function PageFilesSize {
    $pages=Get-CimInstance Win32_PageFile | Select-Object Name, InitialSize, MaximumSize, Filesize
    Write-Host "Pagefiles:"
    Write-Host ( $pages | Format-Table | Out-String)
}
function ProfileGame {
    $memory = Get-CIMInstance Win32_OperatingSystem | Select FreePhysicalMemory,TotalVisibleMemory
    $gameMemory = Get-Process $gameName | Select-Object Name,@{Name='WorkingSet';Expression={($_.WorkingSet/1KB)}}
    $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    Write-Host "Resources at $timeStamp"
    Write-Host "Memory Info:"
    Write-Host ($memory| Format-Table | Out-String)
    Write-Host "Game memory usage: $gameMemory" 
    DrivesFreeSpace
    PageFilesSize
}


if($debug){
    Write-Host ""
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

Write-Host "Script Started at $timestamp"

# Open bugtracking
if($openBugtracking){ Start-Process $bugtracking}

# Launch OBS
$obsActive = Get-Process obs64 -ErrorAction SilentlyContinue
if($obsActive -eq $null){
 $obs = Start-Process "$obsWd\obs64.exe" -PassThru -WorkingDirectory $obsWd
 }else{
 $obs = Get-Process obs64
 }
$obsId = ($obs.Id)
If ($debug) { Write-Host "OBS PID: $obsId"}

# Launch LibreOffice Writer
$writerActive = Get-Process $libreProcess -ErrorAction SilentlyContinue
if($writerActive -eq $null){ $writer = Start-Process $librePath -PassThru}
else{
 $writer = Get-Process $libreProcess
 }
$writerId = ($writer.Id)
If ($debug) { Write-Host "Writer PID: $writerId" }

#Launch RamMap
$rmapActive = Get-Process RAMMap64 -ErrorAction SilentlyContinue
if($rmapActive -eq $null) { $rmap = Start-Process $rmapPath -PassThru}
else{
 $rmap = Get-Process RAMMap64
 }
$rmapId = ($rmap.Id)
If ($debug) { Write-Host "RamMap PID: $rmapId" }

# Launch Process Explorer
$procexpActive = Get-Process procexp64 -ErrorAction SilentlyContinue
if($procexpActive -eq $null){ $procexp = Start-Process $procexpPath -PassThru}
else{
 $procexp = Get-Process procexp64
 }
$procexpId = ($procexp.Id)
If ($debug) { Write-Host "Procexp PID: $procexpId" }

# Launch notepad for notes
$notepadActive = Get-Process $notepadProcess -ErrorAction SilentlyContinue
If($notepadActive -eq $null){
If(Test-Path -Path $notes -PathType Leaf){
    $notepad = Start-Process $notepadPath -PassThru -ArgumentList "$notes"
}else{$notepad = Start-Process $notepadPath -PassThru}
}
else{
 $notepad = Get-Process $notepadProcess
 }
$notepadId = ($notepad.Id)
If ($debug) { Write-Host "Notepad PID: $notepadId"}

# Launch Screenshot app
$screenshotActive = Get-Process $screenshotAppProcessName -ErrorAction SilentlyContinue
if($screenshotActive -eq $null) { $screenshot = Start-Process $screenshotAppPath -PassThru}
else{
 $screenshot = Get-Process $screenshotAppProcessName
 }
$screenshotId = ($screenshot.Id)
If ($debug) { Write-Host "Screenshot app PID: $screenshotId"}

# Return to script working directory
Set-Location -Path $wd

# Print resources at start
DrivesFreeSpace
PageFilesSize

#Try to launch the game
if($gameLaunch){Start-Process $gamePath}

# Wait for game to start
$started = $false
Do {
$game = Get-Process $gameName -ErrorAction SilentlyContinue
   If (!($game)) { Write-Host 'Waiting for game to start' ; Start-Sleep -Seconds $gameLaunchWaitTime }
   Else{ 
   $gameId = (Get-Process $gameName).Id
   Write-Host "Game running, pid: $gameId"
   $started = $true 
   }
}
Until($started)
$game = (Get-Process $gameName).Id
If($debug) { Write-Host "Game PID: $game"}

#Launch VMMap
$vmapActive = Get-Process vmmap64 -ErrorAction SilentlyContinue
if($vmapActive) { $vmapActive | Stop-Process -Force }
$vmap = Start-Process "$vmapPath" -PassThru -ArgumentList "-p $game"
$vmapId = ($vmap.Id)
If ($debug) { Write-Host "VmmMap PID: $vmapId" }

# Profiling Loop
$gameRunning = (Get-Process $gameName -ErrorAction SilentlyContinue)
Do {
    #Clear-Host
    ProfileGame
    Start-Sleep -Seconds $profilerRefreshRate
    $gameRunning = (Get-Process $gameName -ErrorAction SilentlyContinue)
}
While($gameRunning)

#Resources after game stops running
$timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
Write-Host "[$timestamp] Game stopped running!"
DrivesFreeSpace
PageFilesSize

