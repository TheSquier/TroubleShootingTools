
<#
############################ Synopsis ###################################
Automating the basic troubleshooting solutions we employ on a daily
basis in tier 1

#################### Description and Features ###########################

This script using PowerShell to automate manual troubleshooting.
It can solve the following issues:
Monitor flickering, Internet Connectivity, Sound issues, Avaya Issues,
Cache issues.

It does this by employing these tactics:
Updating drivers (Audio and Display), Disabling IPv6, Firewall, 
flushing DNS, reseting avaya profiles, and clearing caches.

This script also features a diagnostic mode which can be used to detect
problems. This mode tests the following things:
Internet, Recent BSODs, and windows event viewer critical events.


######################## Updated Versions ##############################

Please check the DevOps Repo for updated versions. 

########################## Author Info #################################

Author: Alex Squier
Personal Email: xandersquier@outlook.com
LinkedIn: 
GitHub:

413*
########################################################################
#>
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force 

Function Driver-Updater {
#updates drivers
clear
$UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
$UpdateSvc.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"")
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher() 
 
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party
          
$Criteria = "IsInstalled=0 and Type='Driver'"
Write-Host('Searching Driver-Updates...') -Fore Green     
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates
if([string]::IsNullOrEmpty($Updates)){
  Write-Host "No pending driver updates."
}
else{
  #Show available Drivers...
  $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl
  $UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
  $updates | % { $UpdatesToDownload.Add($_) | out-null }
  Write-Host('Downloading Drivers...')  -Fore Green
  $UpdateSession = New-Object -Com Microsoft.Update.Session
  $Downloader = $UpdateSession.CreateUpdateDownloader()
  $Downloader.Updates = $UpdatesToDownload
  $Downloader.Download()
  $UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
  $updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }
 
  Write-Host('Installing Drivers...')  -Fore Green
  $Installer = $UpdateSession.CreateUpdateInstaller()
  $Installer.Updates = $UpdatesToInstall
  $InstallationResult = $Installer.Install()
  if($InstallationResult.RebootRequired) { 
  Write-Host('Reboot required! Please reboot now.') -Fore Red
  } else { Write-Host('Done.') -Fore Green }
  $updateSvc.Services | ? { $_.IsDefaultAUService -eq $false -and $_.ServiceID -eq "7971f918-a847-4430-9279-4a52d1efe18d" } | % { $UpdateSvc.RemoveService($_.ServiceID) }
}}
Function Extend-Displays {
Write-Host "Extending Displays.."
start-sleep -s 1
C:\Windows\System32\DisplaySwitch.exe /extend
}
Function Get-Audio-Device {
#gets current audio devices
clear-host
Add-Type @'
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
    int a(); int o();
    int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int f();
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

public static string GetDefault (int direction) {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(direction, 1, out dev));
    string id = null;
    Marshal.ThrowExceptionForHR(dev.GetId(out id));
    return id;
}
'@ -name audio -Namespace system

function getFriendlyName($id) {
    $reg = "HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\MMDEVAPI\$id"
    return (get-ItemProperty $reg).FriendlyName
}

$id0 = [audio]::GetDefault(0)
$id1 = [audio]::GetDefault(1)
write-host "Default Speaker: $(getFriendlyName $id0)" 
write-host "Default Micro  : $(getFriendlyName $id1)"
}
Function Write-Log {
#creates log file
   [CmdletBinding()]

   Param(

 

   [Parameter(Mandatory=$True)]

   [string]

   $Message,

 

   [Parameter(Mandatory=$False)]

   [string]

   $logfile

   )

 

   $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

   $Line = "$Stamp $Message"

   If($logfile) {

       Add-Content $logfile -Value $Line

   }

   Else {

       Write-Output $Line

   }

}
Function Internet-Test {
#tests internet connection
    $ErrorNumber = 0
    clear
    Write-Host "testing Internet"
    Write-Host "Domain Connection test complete"       
    Write-Host "Testing Internet Connection..." #Tests ability to connect to the internet
    try {
        Test-Connection 8.8.8.8 -ErrorAction stop |Out-Null
        write-host "Internet Connection Successful"
        Write-Log "Successfully Contacted Internet" $logfile}
    Catch {
    Write-Host "Failed to connect to internet"
    Write-Log "!!Failed to connect to internet" $logfile
    $ErrorNumber = $ErrorNumber + 1}
    Write-Host "Internet connection tested."
    Write-Host "Testing DNS..." #tests ability to resolve DNS names
     try {
        Test-Connection google.com -ErrorAction stop | Out-Null
        Write-Host "DNS Successful"
        Write-Log "Successfully resolved DNS" $logfile
        }
    Catch {
    Write-Host "Failed to resolve DNS"
    Write-Log "!!Failed to resolve DNS" $logfile
    $ErrorNumber = $ErrorNumber + 1}
    Write-Host "DNS Tested."
    Write-Host "Internet Testing Completed. $($Errorcounter) Errors found."
    }
function Monitors {
clear
Write-Host "troubleshooting Monitors..."
start-sleep -s 1
Driver-Updater
Write-Log "Drivers Updated" $logfile
}
Function Internet-TS {
#Troubleshooting internet connectivity by disabling the firewall, disabling IPv6, and flushing DNS
Write-Host "Begining Internet Troubleshooter..."
Start-Sleep -s 2 
Write-Host "disabling Firewall..."
start-sleep -s 2
Start-Process powershell -Verb RunAs {netsh advfirewall set allprofiles state off}
Write-Host "Disabling IPv6..."
Start-Sleep -s 2
Start-Process powershell -Verb RunAs {Disable-NetAdapterBinding -Name * -ComponentID 'ms_tcpip6' -Confirm:$false -erroraction 'silentlycontinue'} #Disable IPV6 adapters.
ipconfig /flushdns
}
Function Audio-TS {
#Troubleshoots audio devices by updating drivers.
    Write-Host "Updating Drivers..."
    Start-sleep -s 2
    Driver-Updater
    Write-Host "Pulling current default devices. If incorrect, go to sound settings and change."
    Start-Sleep -s 2
    Get-Audio-Device
    }
function Clear-Caches {
#clears browser caches
clear
Write-Host "Clearing Chrome Cache..."
Remove-Item -path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache2\entries\*" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -EA SilentlyContinue -Verbose
Write-Host "Clearing edge Cache..."
Remove-Item -path "$env:localappdata\Microsoft\Edge\User Data\Default\Cache\*" -Recurse -Force -EA SilentlyContinue -Verbose
Write-Host "Clearing Teams Cache..."

}
Function BSOD-Scan {
try {
start-sleep -s 1
Get-EventLog -LogName application -Newest 1000 -Source ‘Windows Error*' | select timewritten, message | where message -match ‘bluescreen’ |  ft -auto -wrap
$ErrorNumber = $ErrorNumber + 1
$player = New-Object System.Media.SoundPlayer "C:\WINDOWS\Media\Windows Error.wav"
$player.PlayLooping()
start-sleep -s 1
$player.Stop()}
Catch {
Write-Host "No BSOD Logs found."}
}
Function Diagnostic-Mode {
#A diagnostic tool which collects BSOD, Critical Events and other testing and stores as a log file.

Write-Host "Running Diagnostics..."
Write-Host "Running internet test..."

Internet-Test
pause
Write-host "Scanning for recent BSOD..."
start-sleep -s 3
BSOD-Scan
BSOD-Scan | Out-File "<log file>"
Get-WinEvent -FilterHashTable @{logname="system"; Level=1} | Out-File "\\geha\share\Media Library\ServiceDesk\Scripts\AutoTroubleShooter Logs\$($clockID)-$($date)\CriticalEvents.txt"
function Start-Menu { 
#this displays the text for the start up menu
    Clear-Host
    Write-Host "
============ Auto Troubleshooter v 1.1 ===============

                        413*

     Created by Alex Squier last updated 04/23     
                     Loading... 
=====================================================


"
    start-sleep -s 2
    Write-Host "1: Monitors, Displays, and Screens (Updates Monitor 
drivers)
"
    Write-Host "2: Internet, Connectivity (Disables Firewall, Disables 
IPv6 then flushes DNS)
"
    Write-Host "3: Sound (Prints current sound output/input then
updates drivers
"
    Write-Host "5: Caches(clears caches for teams, chrome and edge)
    "
    Write-Host "6: Diagnostics Mode (Check internet, and recent BSODs)

"
    Write-Host "
Q: Press 'Q' to quit.
"
}
Function Speed-Up {
Write-Host "Flushing DNS..."
ipconfig /flushdns
}
#setup and variable initialization
mkdir "<log file>"
$clockID = $env:UserName #gets clock ID
$date = Get-Date -Format "MM-dd-yyyy" #gets date for a variable
$logfile = "<log file>"
$Request = " "
$errorcounter = 0
Write-Log "Clock ID = $($ClockID)" $logfile
Write-Log "Computer Name: $($env:computername)" $logfile

do
{
#This processes menu options and starts the relevant script
    Start-Menu 
    $selection = Read-Host "Please make a selection. Enter a number and press Enter."
    switch ($selection)
    { '1' {
    Write-Log "Running monitor troubleshooter" $logfile
    Monitors
    } '2' {
    Write-Log "Running internet troubleshooter" $logfile
    Internet-TS
    internet-test
    } '3' {
    Write-Log "Running audio troubleshooter" $logfile
    Get-Audio-Device
    pause
    Write-Log "Updating drivers" $logfile
    Driver-Updater
    } '4' {
    Write-Log "Running avaya reset" $logfile
    Avaya-Reset
    } '5' {
    Write-Log "Clearing caches" $logfile
    Clear-Caches
    } '6' {
    Write-Log "Running Diagnostics mode" $logfile
    Diagnostic-Mode
    pause
    }}
}
until ($selection -eq 'q')