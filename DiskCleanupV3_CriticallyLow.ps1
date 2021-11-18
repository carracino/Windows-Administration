#  Drive cleanup script. Critically Low Disk Space machines only* 
# Version 3.0 : Enterprise cleanup. Target = all windows devices with less than 10GB free. 
# This version is to free space and remove identified items that are not required to remain on disk and will improve performance.  
# Author: Jon Carracino 4-10-19

#ensure 64Bit PS runtime:
if (($pshome -like "*syswow64*") -and ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -like "64*")) 
{
    write-warning "Restarting script under 64 bit powershell"
 
    # relaunch this script under 64 bit shell
    & (join-path ($pshome -replace "syswow64", "sysnative")\powershell.exe) -file $myinvocation.mycommand.Definition @args
 
    # This will exit the original powershell process. This will only be done in case of an x86 process on a x64 OS.
    exit
}


# variables:
# below locations will be targeted for removal of old files based on Daysback setting below:
$Path1 = $env:TEMP
$Path2 = "C:\Users\*\Appdata\Local\Temp"
$Path3 = "C:\Users\*\Appdata\Local\Microsoft\Windows\Temporary Internet Files"
$path4 = "C:\Users\*\Appdata\Local\Google\Chrome\User Data\Default\Cache"
$path5 = "C:\Users\*\Appdata\Local\Google\Chrome\User Data\Profile 1\Cache"
$path6 = "C:\Users\*\Appdata\Local\Google\Chrome\User Data\Profile 2\Cache"
$path7 = "$env:windir\minidump"
#$path8 = "$env:windir\Prefetch"

$Daysback = "-3"
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($Daysback)
$logpath = "$env:windir\installer\DiskCleanup.log"
$Disk1 = gwmi Win32_LogicalDisk -Property * -Filter "DeviceID='C:'"

### Log Function
#.EXAMPLE 
#   Write-Log -Message 'Folder does not exist.' -Path "any path other than default" -Level Info 
function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path=$logpath, 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}
#####  End Function

#Start log
Write-Log -Message "$Date : Begin Cleanup ###################################" -Level Info

#log free space at start:
$freespace = $Disk1.Freespace / 1GB
Write-Log -Message "FreeSpace before cleanup: $freespace  GB  " -Level Info

# Delete all Files in System temp older than 5 day(s) 
Get-ChildItem $Path1 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# Delete all user temp files:
Get-ChildItem $Path2 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# Delete all profile temp internet content - includes office cache.
Get-ChildItem $Path3 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# Delete all profile Chrome cache.
Get-ChildItem $Path4 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Get-ChildItem $Path5 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Get-ChildItem $Path6 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
#Remove mini dump files
Get-ChildItem $Path7 -Recurse -Force | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

#Remove nomad cache older than 5 days (only for critical low machines):
Start-Process cachecleaner.exe -ArgumentList "-maxcacheage=5" -wait -NoNewWindow -PassThru

#run disk clean util for Win10:

#Setup specific items to clean:
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
#only do update cleanup and DO cleanup on critically low space devices
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Delivery Optimization Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue
New-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache" -Name "StateFlags0100" -value "2" -PropertyType DWORD -ErrorAction SilentlyContinue

Start-Process cleanmgr.exe -ArgumentList "/sagerun:100" -wait -NoNewWindow -PassThru
#Only for critically low space machines:
    Start-Process cleanmgr.exe -ArgumentList "/AUTOCLEAN" -wait -NoNewWindow -PassThru
    #Turn off hibernation feature:
    Start-Process "cmd.exe" -ArgumentList "/c powercfg.exe -h off" -wait -NoNewWindow -PassThru

#log free space at end:
$Disk1 = gwmi Win32_LogicalDisk -Property * -Filter "DeviceID='C:'"
$freespace = $Disk1.Freespace / 1GB
Write-Log -Message "FreeSpace after cleanup: $freespace  GB  " -Level Info

Return 0

#END