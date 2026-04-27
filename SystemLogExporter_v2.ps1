
#Requires -Version 5.1
# Author: Jon Carracino
# Version 1.0 - 04-2026
<#
.SYNOPSIS
    SystemLogExporter -- Windows System Log Export Tool
.DESCRIPTION
    WPF-based GUI for comprehensive local Windows system log exporting.
    Exports findings as CSV files to a timestamped report folder.
.NOTES
    Run as Administrator for complete results (security events, all services, etc.)
    Author : SystemLogExporter
    Version: 1.0
#>

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'

# ── Check for admin elevation and prompt if needed ────────────────────────────
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$script:IsAdmin   = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $script:IsAdmin) {
    Add-Type -AssemblyName PresentationFramework
    $answer = [System.Windows.MessageBox]::Show(
        "This script is not running with Administrator privileges.`n`n" +
        "Some modules (Security Event Logs, Group Policy, Windows Update Logs) require elevation.`n`n" +
        "Would you like to relaunch as Administrator?",
        "SystemLogExporter - Elevation Required",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
        try {
            $psExe = (Get-Process -Id $PID).Path
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName  = $psExe
            $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            $startInfo.Verb      = 'RunAs'
            [System.Diagnostics.Process]::Start($startInfo) | Out-Null
            exit
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Failed to elevate: $($_.Exception.Message)`n`nContinuing without admin privileges.",
                "SystemLogExporter",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }
    # If user said No or elevation failed, re-check in case elevation happened
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $script:IsAdmin   = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

#region ── XAML ────────────────────────────────────────────────────────────────
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SystemLogExporter"
    Height="720" Width="960"
    MinHeight="560" MinWidth="760"
    WindowStartupLocation="CenterScreen"
    Background="#0F1117"
    Foreground="#E2E8F0"
    FontFamily="Segoe UI">

    <Window.Resources>

        <!-- Scrollbar -->
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="6"/>
            <Setter Property="Background" Value="Transparent"/>
        </Style>

        <!-- Primary button -->
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="18,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#60A5FA"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#1E2435"/>
                                <Setter Property="Foreground" Value="#475569"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Ghost button -->
        <Style x:Key="BtnGhost" TargetType="Button">
            <Setter Property="Background" Value="#1A1F2E"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#2D3748"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#252D40"/>
                                <Setter TargetName="bd" Property="BorderBrush" Value="#3B82F6"/>
                                <Setter Property="Foreground" Value="#E2E8F0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="#334155"/>
                                <Setter TargetName="bd" Property="BorderBrush" Value="#1E2435"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Input -->
        <Style x:Key="Input" TargetType="TextBox">
            <Setter Property="Background" Value="#151A27"/>
            <Setter Property="Foreground" Value="#CBD5E1"/>
            <Setter Property="BorderBrush" Value="#2D3748"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="CaretBrush" Value="#3B82F6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Label -->
        <Style x:Key="Lbl" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#64748B"/>
            <Setter Property="FontSize" Value="11.5"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

    </Window.Resources>

    <Grid>
        <!-- Accent sidebar stripe -->
        <Border Width="3" HorizontalAlignment="Left" Background="#3B82F6" Opacity="0.6"/>

        <Grid Margin="24,20,20,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- ── Header ── -->
            <Grid Grid.Row="0" Margin="0,0,0,22">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="SYSTEMLOGEXPORTER" FontSize="20" FontWeight="Bold"
                               Foreground="#F1F5F9"
                               FontFamily="Consolas"/>
                    <TextBlock Text="Windows System Log Export Tool  •  Local Machine"
                               FontSize="11" Foreground="#475569" Margin="2,3,0,0"/>
                </StackPanel>
                <Border Grid.Column="1" x:Name="ElevationBadge"
                        Background="#1A2E1A" BorderBrush="#166534" BorderThickness="1"
                        CornerRadius="4" Padding="10,4">
                    <TextBlock x:Name="ElevationText" Text="● ELEVATED" FontSize="10.5"
                               FontWeight="SemiBold" Foreground="#4ADE80" FontFamily="Consolas"/>
                </Border>
            </Grid>

            <!-- ── Output Folder ── -->
            <Grid Grid.Row="1" Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Output Folder" Style="{StaticResource Lbl}"/>
                <TextBox   Grid.Column="1" x:Name="TxtOutput" Style="{StaticResource Input}"
                           IsReadOnly="True" VerticalAlignment="Center"/>
                <Button    Grid.Column="2" x:Name="BtnBrowse" Content="Browse..."
                           Style="{StaticResource BtnGhost}" Margin="8,0,0,0" VerticalAlignment="Center"/>
            </Grid>

            <!-- ── Look-back Days ── -->
            <Grid Grid.Row="2" Margin="0,0,0,18">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="80"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Look-back Days" Style="{StaticResource Lbl}"/>
                <TextBox   Grid.Column="1" x:Name="TxtDays" Style="{StaticResource Input}"
                           Text="30" TextAlignment="Center" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="2" Margin="12,0,0,0" Style="{StaticResource Lbl}"
                           Text="Applied to event logs, recently modified files, and crash history"/>
            </Grid>

            <!-- ── Divider ── -->
            <Border Grid.Row="3" Height="1" Background="#1A2035" Margin="0,0,0,14"/>

            <!-- ── Live Log ── -->
            <Border Grid.Row="4" Background="#0C1020" BorderBrush="#1E2A40"
                    BorderThickness="1" CornerRadius="6" Margin="0,0,0,12">
                <ScrollViewer x:Name="LogScroller" VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Auto" Padding="4">
                    <TextBox x:Name="TxtLog"
                             Background="Transparent" Foreground="#94A3B8"
                             BorderThickness="0" IsReadOnly="True"
                             TextWrapping="NoWrap"
                             FontFamily="Consolas" FontSize="11.5"
                             Padding="10,8" VerticalAlignment="Top"
                             AcceptsReturn="True"/>
                </ScrollViewer>
            </Border>

            <!-- ── Progress Bar ── -->
            <ProgressBar Grid.Row="5" x:Name="PBar" Height="4"
                         Minimum="0" Maximum="100" Value="0"
                         Foreground="#3B82F6" Background="#1A1F2E"
                         BorderThickness="0" Margin="0,0,0,14"/>

            <!-- ── Footer Actions ── -->
            <Grid Grid.Row="6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtStatus" Grid.Column="0"
                           Text="Select an output folder and click Run Export."
                           Style="{StaticResource Lbl}" FontSize="11"/>
                <Button Grid.Column="1" x:Name="BtnOpen" Content="Open Report Folder"
                        Style="{StaticResource BtnGhost}" Margin="0,0,8,0"
                        IsEnabled="False" Visibility="Collapsed"/>
                <Button Grid.Column="2" x:Name="BtnRun" Content="▶  Run Export"
                        Style="{StaticResource BtnPrimary}" IsEnabled="False"/>
            </Grid>

        </Grid>
    </Grid>
</Window>
'@
#endregion

#region ── Load UI ────────────────────────────────────────────────────────────
$reader  = [System.Xml.XmlNodeReader]::new($xaml)
$window  = [System.Windows.Markup.XamlReader]::Load($reader)

$TxtOutput       = $window.FindName('TxtOutput')
$TxtDays         = $window.FindName('TxtDays')
$BtnBrowse       = $window.FindName('BtnBrowse')
$TxtLog          = $window.FindName('TxtLog')
$LogScroller     = $window.FindName('LogScroller')
$PBar            = $window.FindName('PBar')
$TxtStatus       = $window.FindName('TxtStatus')
$BtnRun          = $window.FindName('BtnRun')
$BtnOpen         = $window.FindName('BtnOpen')
$ElevationBadge  = $window.FindName('ElevationBadge')
$ElevationText   = $window.FindName('ElevationText')
#endregion

#region ── Elevation Badge ────────────────────────────────────────────────────
if (-not $script:IsAdmin) {
    $ElevationText.Text = "* NOT ELEVATED"
    $ElevationText.Foreground = [System.Windows.Media.Brushes]::IndianRed
    $ElevationBadge.BorderBrush = [System.Windows.Media.SolidColorBrush][System.Windows.Media.Color]::FromRgb(0x7F,0x1D,0x1D)
    $ElevationBadge.Background  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.Color]::FromRgb(0x1F,0x10,0x10)
}
#endregion

#region ── Sync Hash ──────────────────────────────────────────────────────────
$syncHash = [hashtable]::Synchronized(@{
    Log        = $TxtLog
    Scroller   = $LogScroller
    PBar       = $PBar
    Status     = $TxtStatus
    RunBtn     = $BtnRun
    OpenBtn    = $BtnOpen
    Dispatcher = $window.Dispatcher
    ReportPath = ''
    IsAdmin    = $script:IsAdmin
})
#endregion

#region ── Browse ─────────────────────────────────────────────────────────────
$BtnBrowse.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description  = 'Select output folder for audit reports'
    $dlg.SelectedPath = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtOutput.Text      = $dlg.SelectedPath
        $BtnRun.IsEnabled    = $true
        $TxtStatus.Text      = "Ready -- click Run Export to begin."
    }
})
#endregion

#region ── Open Folder ────────────────────────────────────────────────────────
$BtnOpen.Add_Click({
    $p = $syncHash.ReportPath
    if ($p -and (Test-Path $p)) { Start-Process explorer.exe "`"$p`"" }
})
#endregion

#region ── Audit Script (runs in background runspace) ─────────────────────────
$auditScript = {
    param([string]$RootOutput, [int]$DaysBack, [hashtable]$SH)

    Set-StrictMode -Off
    $ErrorActionPreference = 'SilentlyContinue'
    $script:SH         = $SH
    $script:IsAdmin    = $SH.IsAdmin
    $script:TotalSteps = 25
    $script:StepsDone  = 0

    # ── Helpers ──────────────────────────────────────────────────────────────
    function Write-Log {
        param([string]$Msg, [switch]$Status)
        $ts  = Get-Date -Format 'HH:mm:ss'
        $line = "[$ts]  $Msg"
        $script:SH.Dispatcher.Invoke([action]{
            $script:SH.Log.AppendText("$line`n")
            $script:SH.Scroller.ScrollToBottom()
            if ($Status) { $script:SH.Status.Text = $Msg }
        }, 'Normal')
    }

    function Step-Progress {
        $script:StepsDone++
        $pct = [math]::Round(($script:StepsDone / $script:TotalSteps) * 100)
        $script:SH.Dispatcher.Invoke([action]{ $script:SH.PBar.Value = $pct }, 'Normal')
    }

    function Save-Module {
        param([string]$FileName, [object]$Data, [string]$OutDir)
        $arr = @($Data)
        if ($arr.Count -gt 0 -and $null -ne $arr[0]) {
            $arr | Export-Csv -Path "$OutDir\$FileName.csv" -NoTypeInformation -Encoding UTF8
            Write-Log "  [OK]  $FileName  ($($arr.Count) records)"
        }
        else {
            Write-Log "  --  $FileName  (no data collected)"
        }
    }

    # ── Create report folder ─────────────────────────────────────────────────
    $stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
    $hostname  = $env:COMPUTERNAME
    $reportDir = Join-Path $RootOutput "SystemLogExporter_${hostname}_${stamp}"
    New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    $script:SH.Dispatcher.Invoke([action]{ $script:SH.ReportPath = $reportDir }, 'Normal')

    $cutoff = (Get-Date).AddDays(-$DaysBack)

    Write-Log "=============================================="
    Write-Log "  SYSTEMLOGEXPORTER  |  $hostname"
    Write-Log "  Output  : $reportDir"
    Write-Log "  RunBy   : $env:USERNAME"
    Write-Log "  Cutoff  : $($cutoff.ToString('yyyy-MM-dd'))  ($DaysBack days back)"
    Write-Log "=============================================="

    $summary = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 01 — System Information + Installed Patches
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [01/25]  System Information & Installed Patches"
    try {
        $os   = Get-CimInstance Win32_OperatingSystem
        $cs   = Get-CimInstance Win32_ComputerSystem
        $bios = Get-CimInstance Win32_BIOS
        $cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
        $up   = (Get-Date) - $os.LastBootUpTime

        $sysRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $add = { param($k,$v) $sysRows.Add([PSCustomObject]@{ Property=$k; Value=$v }) }

        & $add 'ComputerName'       $env:COMPUTERNAME
        & $add 'Domain'             $cs.Domain
        & $add 'OS'                 $os.Caption
        & $add 'OSVersion'          $os.Version
        & $add 'BuildNumber'        $os.BuildNumber
        & $add 'Architecture'       $os.OSArchitecture
        & $add 'InstallDate'        $os.InstallDate
        & $add 'LastBoot'           $os.LastBootUpTime
        & $add 'Uptime'             "$($up.Days)d $($up.Hours)h $($up.Minutes)m"
        & $add 'Manufacturer'       $cs.Manufacturer
        & $add 'Model'              $cs.Model
        & $add 'TotalRAM_GB'        $([math]::Round($cs.TotalPhysicalMemory/1GB, 2))
        & $add 'CPU'                $cpu.Name
        & $add 'CPUCores'           $cpu.NumberOfCores
        & $add 'CPULogicalProc'     $cpu.NumberOfLogicalProcessors
        & $add 'BIOSVersion'        $bios.SMBIOSBIOSVersion
        & $add 'BIOSReleaseDate'    $bios.ReleaseDate
        & $add 'SerialNumber'       $bios.SerialNumber
        & $add 'TimeZone'           (Get-TimeZone).DisplayName
        & $add 'AuditBy'            $env:USERNAME
        & $add 'AuditDateTime'      (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

        Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
            & $add "Disk_$($_.DeviceID)" $("Total: $([math]::Round($_.Size/1GB,1))GB | Free: $([math]::Round($_.FreeSpace/1GB,1))GB")
        }

        $patches = Get-HotFix | Select-Object HotFixID, Description, InstalledOn, InstalledBy | Sort-Object InstalledOn -Descending

        Save-Module '01_SystemInfo'       $sysRows  $reportDir
        Save-Module '02_InstalledPatches' $patches  $reportDir
        $summary.Add([PSCustomObject]@{ Module='SystemInfo';     Status='OK'; Records=$sysRows.Count })
    }
    catch { Write-Log "  [!]  System Info ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='SystemInfo'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 02 — User Accounts & Group Memberships
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [02/25]  User Accounts & Group Memberships"
    try {
        $users = Get-LocalUser | Select-Object Name, Enabled, LastLogon,
            PasswordRequired, PasswordExpires, PasswordLastSet,
            Description, SID,
            @{N='AccountExpires';E={ if ($_.AccountExpires) {$_.AccountExpires} else {'Never'} }}

        $groupRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        Get-LocalGroup | ForEach-Object {
            $grp     = $_
            $members = try { Get-LocalGroupMember -Group $grp.Name -EA Stop } catch { @() }
            if ($members) {
                $members | ForEach-Object {
                    $groupRows.Add([PSCustomObject]@{
                        GroupName   = $grp.Name
                        Description = $grp.Description
                        MemberName  = $_.Name
                        MemberType  = $_.ObjectClass
                        MemberSID   = $_.SID
                    })
                }
            } else {
                $groupRows.Add([PSCustomObject]@{
                    GroupName=$grp.Name; Description=$grp.Description
                    MemberName="$($grp.Name) (empty)"; MemberType=''; MemberSID=''
                })
            }
        }

        Save-Module '03_UserAccounts'    $users              $reportDir
        Save-Module '04_GroupMemberships' $groupRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='UserAccounts'; Status='OK'; Records=($users | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  User Accounts ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='UserAccounts'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 03 — Running Processes & Services
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [03/25]  Running Processes & Services"
    try {
        $procs = Get-CimInstance Win32_Process | Select-Object ProcessId, Name, ExecutablePath,
            CommandLine, ParentProcessId, CreationDate,
            @{N='Owner';E={ ($_ | Invoke-CimMethod -MethodName GetOwner -EA SilentlyContinue).User }} |
            Sort-Object Name

        $services = Get-CimInstance Win32_Service | Select-Object Name, DisplayName, State,
            StartMode, StartName, PathName, Description | Sort-Object State, Name

        Save-Module '05_Processes' $procs    $reportDir
        Save-Module '06_Services'  $services $reportDir
        $summary.Add([PSCustomObject]@{ Module='Processes'; Status='OK'; Records=($procs | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Processes/Services ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='Processes'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 04 — Scheduled Tasks
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [04/25]  Scheduled Tasks"
    try {
        $tasks = Get-ScheduledTask | ForEach-Object {
            $info = $_ | Get-ScheduledTaskInfo -EA SilentlyContinue
            [PSCustomObject]@{
                TaskName    = $_.TaskName
                TaskPath    = $_.TaskPath
                State       = $_.State
                Author      = $_.Principal.UserId
                RunLevel    = $_.Principal.RunLevel
                LastRunTime = $info.LastRunTime
                NextRunTime = $info.NextRunTime
                LastResult  = $info.LastTaskResult
                Actions     = ($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)".Trim() }) -join ' | '
                Triggers    = ($_.Triggers | ForEach-Object { $_.CimClass.CimClassName }) -join ', '
            }
        } | Sort-Object State, TaskPath

        Save-Module '07_ScheduledTasks' $tasks $reportDir
        $summary.Add([PSCustomObject]@{ Module='ScheduledTasks'; Status='OK'; Records=($tasks | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Scheduled Tasks ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='ScheduledTasks'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 05 — Network Configuration
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [05/25]  Network Configuration"
    try {
        $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, Status,
            MacAddress, LinkSpeed,
            @{N='IPAddresses';E={ (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -EA SilentlyContinue | Select-Object -ExpandProperty IPAddress) -join ', ' }},
            @{N='DNSServers'; E={ (Get-DnsClientServerAddress -InterfaceIndex $_.InterfaceIndex -EA SilentlyContinue | Select-Object -ExpandProperty ServerAddresses) -join ', ' }}

        $connections = Get-NetTCPConnection | ForEach-Object {
            [PSCustomObject]@{
                LocalAddress  = $_.LocalAddress
                LocalPort     = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort    = $_.RemotePort
                State         = $_.State
                ProcessId     = $_.OwningProcess
                ProcessName   = (Get-Process -Id $_.OwningProcess -EA SilentlyContinue).Name
            }
        } | Sort-Object State, LocalPort

        $dnsCache = Get-DnsClientCache | Select-Object Entry, RecordName, RecordType,
            Status, TimeToLive, Data

        $shares = Get-SmbShare | Select-Object Name, Path, Description, CurrentUsers,
            EncryptData, FolderEnumerationMode

        $routes = Get-NetRoute -AddressFamily IPv4 | Select-Object DestinationPrefix,
            NextHop, RouteMetric, InterfaceAlias, InterfaceIndex | Sort-Object RouteMetric

        Save-Module '08_NetworkAdapters'  $adapters    $reportDir
        Save-Module '09_TCPConnections'   $connections $reportDir
        Save-Module '10_DNSCache'         $dnsCache    $reportDir
        Save-Module '11_NetworkShares'    $shares      $reportDir
        Save-Module '11b_RoutingTable'    $routes      $reportDir
        $summary.Add([PSCustomObject]@{ Module='Network'; Status='OK'; Records=($connections | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Network ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='Network'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 06 — Startup & Persistence
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [06/25]  Startup & Persistence (Registry Run Keys)"
    try {
        $persist = [System.Collections.Generic.List[PSCustomObject]]::new()

        $runKeys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
            'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon',
            'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\BootExecute'
        )
        foreach ($key in $runKeys) {
            try {
                $props = Get-ItemProperty -Path $key -EA Stop
                $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                    $persist.Add([PSCustomObject]@{ Source='RegistryRunKey'; KeyPath=$key; Name=$_.Name; Value=$_.Value })
                }
            } catch {}
        }

        # Startup folders
        @(
            "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup",
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
        ) | Where-Object { Test-Path $_ } | ForEach-Object {
            Get-ChildItem $_ -File -EA SilentlyContinue | ForEach-Object {
                $persist.Add([PSCustomObject]@{ Source='StartupFolder'; KeyPath=$_; Name=$_.Name; Value=$_.FullName })
            }
        }

        # Auto-start services with non-standard paths (potential lateral movement)
        Get-CimInstance Win32_Service | Where-Object {
            $_.StartMode -eq 'Auto' -and
            $_.PathName  -notmatch 'System32|SysWOW64|Program Files|Windows\\|MsMpEng'
        } | ForEach-Object {
            $persist.Add([PSCustomObject]@{ Source='AutoService_NonStdPath'; KeyPath='SCM'; Name=$_.Name; Value=$_.PathName })
        }

        Save-Module '12_StartupPersistence' $persist.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='Persistence'; Status='OK'; Records=$persist.Count })
    }
    catch { Write-Log "  [!]  Startup/Persistence ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='Persistence'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 07 — Installed Software
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [07/25]  Installed Software"
    try {
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        $software = $regPaths | ForEach-Object {
            Get-ItemProperty $_ -EA SilentlyContinue
        } | Where-Object { $_.DisplayName } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate,
                InstallLocation, UninstallString,
                @{N='Arch';E={ if ($_.PSPath -match 'WOW6432') {'x86'} else {'x64'} }} |
            Sort-Object Publisher, DisplayName

        Save-Module '13_InstalledSoftware' $software $reportDir
        $summary.Add([PSCustomObject]@{ Module='InstalledSoftware'; Status='OK'; Records=($software | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Installed Software ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='InstalledSoftware'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 08 — USB / Removable Device History
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [08/25]  USB & Removable Device History"
    try {
        $usbRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $usbBase = 'HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR'
        if (Test-Path $usbBase) {
            Get-ChildItem $usbBase -EA SilentlyContinue | ForEach-Object {
                $class = $_.PSChildName
                Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
                    $inst  = $_.PSChildName
                    $p     = Get-ItemProperty $_.PSPath -EA SilentlyContinue
                    $usbRows.Add([PSCustomObject]@{
                        DeviceClass  = $class
                        InstanceId   = $inst
                        FriendlyName = $p.FriendlyName
                        DeviceDesc   = $p.DeviceDesc
                        Manufacturer = $p.Mfg
                        SerialNumber = ($inst -split '\\' | Select-Object -Last 1)
                        ContainerId  = $p.ContainerID
                    })
                }
            }
        }

        # Also USBSTOR from SetupAPI logs (broader history, including removed devices)
        $setupApiLog = "$env:SystemRoot\INF\setupapi.dev.log"
        if (Test-Path $setupApiLog) {
            $logContent = Get-Content $setupApiLog -Raw -EA SilentlyContinue
            $usbEntries = [regex]::Matches($logContent, 'USBSTOR\\[^\s]+') |
                Select-Object -ExpandProperty Value | Sort-Object -Unique
            foreach ($entry in $usbEntries) {
                if (-not ($usbRows | Where-Object { $_.InstanceId -eq $entry })) {
                    $usbRows.Add([PSCustomObject]@{
                        DeviceClass  = 'USBSTOR (SetupAPI)'
                        InstanceId   = $entry
                        FriendlyName = ''
                        DeviceDesc   = 'Historical (from SetupAPI log)'
                        Manufacturer = ''
                        SerialNumber = ($entry -split '\\' | Select-Object -Last 1)
                        ContainerId  = ''
                    })
                }
            }
        }

        Save-Module '14_USBHistory' $usbRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='USBHistory'; Status='OK'; Records=$usbRows.Count })
    }
    catch { Write-Log "  [!]  USB History ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='USBHistory'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 09 — Event Log Analysis (Logons, Failures, Account Changes)
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [09/25]  Event Log Analysis  ($DaysBack-day window)"
    if (-not $script:IsAdmin) {
        Write-Log "  --  SKIPPED (Security Event Log requires admin privileges)"
        $summary.Add([PSCustomObject]@{ Module='EventLogs'; Status='Skipped (not elevated)'; Records=0 })
    } else {
    try {
        $eventMap = @{
            4624='Successful Logon'; 4625='Failed Logon'; 4634='Logoff'
            4647='User Initiated Logoff'; 4648='Explicit Credentials Logon'
            4672='Special Privilege Logon'; 4720='Account Created'; 4722='Account Enabled'
            4723='Password Change Attempt'; 4724='Password Reset'; 4725='Account Disabled'
            4726='Account Deleted'; 4728='Added to Global Group'; 4732='Added to Local Group'
            4756='Added to Universal Group'; 4768='Kerberos Auth Request'; 4776='NTLM Auth'
        }

        $secEvents = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawEvts   = Get-WinEvent -FilterHashtable @{
            LogName='Security'; Id=$eventMap.Keys; StartTime=$cutoff
        } -EA SilentlyContinue

        foreach ($evt in $rawEvts) {
            $xml  = [xml]$evt.ToXml()
            $data = @{}
            $xml.Event.EventData.Data | ForEach-Object { if ($_.Name) { $data[$_.Name] = $_.'#text' } }
            $secEvents.Add([PSCustomObject]@{
                TimeCreated   = $evt.TimeCreated
                EventId       = $evt.Id
                EventType     = $eventMap[[int]$evt.Id]
                SubjectUser   = $data['SubjectUserName']
                TargetUser    = $data['TargetUserName']
                Domain        = $data['TargetDomainName']
                LogonType     = $data['LogonType']
                WorkStation   = $data['WorkstationName']
                IPAddress     = $data['IpAddress']
                ProcessName   = $data['ProcessName']
                FailureReason = $data['FailureReason']
            })
        }

        # System-level errors and criticals
        $sysErrors = Get-WinEvent -FilterHashtable @{
            LogName='System'; Level=@(1,2); StartTime=$cutoff
        } -EA SilentlyContinue | Select-Object TimeCreated, Id, LevelDisplayName, ProviderName,
            @{N='Message';E={ if ($_.Message) { $m = $_.Message -replace '\s+',' '; $m.Substring(0, [Math]::Min(400, $m.Length)) } else { '' } }}

        Save-Module '15_SecurityEvents' $secEvents.ToArray() $reportDir
        Save-Module '16_SystemErrors'   $sysErrors           $reportDir
        $summary.Add([PSCustomObject]@{ Module='EventLogs'; Status='OK'; Records=$secEvents.Count })
    }
    catch { Write-Log "  [!]  Event Log ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='EventLogs'; Status="Error: $_"; Records=0 }) }
    }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 10 — Recently Modified Files
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [10/25]  Recently Modified Files  ($DaysBack-day window)"
    try {
        $scanRoots  = @($env:USERPROFILE, $env:TEMP, "$env:SystemDrive\Windows\Temp", "$env:SystemDrive\Temp")
        $recentList = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($root in ($scanRoots | Where-Object { Test-Path $_ })) {
            Get-ChildItem -Path $root -Recurse -File -EA SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $cutoff } |
                ForEach-Object {
                    $recentList.Add([PSCustomObject]@{
                        ScanRoot     = $root
                        FullPath     = $_.FullName
                        Name         = $_.Name
                        Extension    = $_.Extension
                        SizeKB       = [math]::Round($_.Length/1KB, 2)
                        Created      = $_.CreationTime
                        LastModified = $_.LastWriteTime
                        LastAccessed = $_.LastAccessTime
                    })
                }
        }

        Save-Module '17_RecentlyModifiedFiles' ($recentList | Sort-Object LastModified -Descending) $reportDir
        $summary.Add([PSCustomObject]@{ Module='RecentFiles'; Status='OK'; Records=$recentList.Count })
    }
    catch { Write-Log "  [!]  Recently Modified Files ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='RecentFiles'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 11 — Application Run History (Prefetch + UserAssist + BAM)
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [11/25]  Application Run History  (Prefetch / UserAssist / BAM)"
    try {
        $runHistory = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Prefetch files
        $pfPath = "$env:SystemRoot\Prefetch"
        if (Test-Path $pfPath) {
            Get-ChildItem $pfPath -Filter '*.pf' -EA SilentlyContinue | ForEach-Object {
                $runHistory.Add([PSCustomObject]@{
                    Source   = 'Prefetch'
                    ExeName  = ($_.Name -replace '-[A-F0-9]{8}\.pf$', '')
                    FullPath = $_.FullName
                    LastRun  = $_.LastWriteTime
                    Created  = $_.CreationTime
                    SizeKB   = [math]::Round($_.Length/1KB,2)
                })
            }
        }

        # UserAssist (ROT13 decoded)
        function ConvertFrom-Rot13([string]$s) {
            -join ($s.ToCharArray() | ForEach-Object {
                $c = [int]$_
                if    ($c -ge 65 -and $c -le 90)  { [char](($c-65+13)%26+65) }
                elseif($c -ge 97 -and $c -le 122) { [char](($c-97+13)%26+97) }
                else  { $_ }
            })
        }
        $uaBase = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\UserAssist'
        if (Test-Path $uaBase) {
            Get-ChildItem $uaBase -EA SilentlyContinue | ForEach-Object {
                $cp = Join-Path $_.PSPath 'Count'
                if (Test-Path $cp) {
                    (Get-ItemProperty $cp -EA SilentlyContinue).PSObject.Properties |
                        Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                            $dec = ConvertFrom-Rot13 $_.Name
                            if ($dec -match '\.(exe|msc|lnk|cpl)') {
                                $runHistory.Add([PSCustomObject]@{
                                    Source='UserAssist'; ExeName=Split-Path $dec -Leaf
                                    FullPath=$dec; LastRun=''; Created=''; SizeKB=''
                                })
                            }
                        }
                }
            }
        }

        # BAM (Background Activity Moderator) — Win10 1709+
        foreach ($bamKey in @('HKLM:\SYSTEM\CurrentControlSet\Services\bam\State\UserSettings',
                               'HKLM:\SYSTEM\CurrentControlSet\Services\bam\UserSettings')) {
            if (Test-Path $bamKey) {
                Get-ChildItem $bamKey -EA SilentlyContinue | ForEach-Object {
                    $sid = $_.PSChildName
                    (Get-ItemProperty $_.PSPath -EA SilentlyContinue).PSObject.Properties |
                        Where-Object { $_.Name -match '\\' -and $_.Name -notmatch '^PS' } |
                        ForEach-Object {
                            $runHistory.Add([PSCustomObject]@{
                                Source="BAM ($sid)"; ExeName=Split-Path $_.Name -Leaf
                                FullPath=$_.Name; LastRun='BAM_binary'; Created=''; SizeKB=''
                            })
                        }
                }
                break
            }
        }

        Save-Module '18_AppRunHistory' $runHistory.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='AppRunHistory'; Status='OK'; Records=$runHistory.Count })
    }
    catch { Write-Log "  [!]  App Run History ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='AppRunHistory'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 12 — Website Activity (Browser History)
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [12/25]  Website Activity  (Chrome / Edge / Firefox / Brave)"
    try {
        $browserRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        $tmpDir      = Join-Path $env:TEMP "FA_BrowserCache"
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

        $browsers = @(
            @{ Name='Chrome';  Base="$env:LOCALAPPDATA\Google\Chrome\User Data";          HistRel='Default\History';  FF=$false }
            @{ Name='Edge';    Base="$env:LOCALAPPDATA\Microsoft\Edge\User Data";          HistRel='Default\History';  FF=$false }
            @{ Name='Brave';   Base="$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"; HistRel='Default\History'; FF=$false }
            @{ Name='Firefox'; Base="$env:APPDATA\Mozilla\Firefox\Profiles";              HistRel='places.sqlite';    FF=$true  }
        )

        foreach ($b in $browsers) {
            if (-not (Test-Path $b.Base)) { continue }

            if ($b.FF) {
                # Firefox: iterate .default profiles
                Get-ChildItem $b.Base -Directory -EA SilentlyContinue |
                    Where-Object { $_.Name -match '\.default' } | ForEach-Object {
                        $profileName = $_.Name
                        $dbSrc = Join-Path $_.FullName 'places.sqlite'
                        if (-not (Test-Path $dbSrc)) { return }
                        $dbTmp = Join-Path $tmpDir "ff_places.db"
                        Copy-Item $dbSrc $dbTmp -Force -EA SilentlyContinue
                        if (-not (Test-Path $dbTmp)) { return }
                        $bytes = [System.IO.File]::ReadAllBytes($dbTmp)
                        $text  = [System.Text.Encoding]::UTF8.GetString($bytes)
                        [regex]::Matches($text, 'https?://[^\x00-\x1F\x7F\s"<>|]{10,250}') |
                            Select-Object -ExpandProperty Value | Sort-Object -Unique | ForEach-Object {
                                $browserRows.Add([PSCustomObject]@{ Browser='Firefox'; Profile=$profileName; URL=$_; Method='BinaryExtract' })
                            }
                        Remove-Item $dbTmp -Force -EA SilentlyContinue
                    }
            }
            else {
                $histSrc = Join-Path $b.Base $b.HistRel
                if (-not (Test-Path $histSrc)) { continue }
                $dbTmp = Join-Path $tmpDir "$($b.Name)_hist.db"
                Copy-Item $histSrc $dbTmp -Force -EA SilentlyContinue
                if (-not (Test-Path $dbTmp)) { continue }
                $bytes = [System.IO.File]::ReadAllBytes($dbTmp)
                $text  = [System.Text.Encoding]::UTF8.GetString($bytes)
                [regex]::Matches($text, 'https?://[^\x00-\x1F\x7F\s"<>|]{10,250}') |
                    Select-Object -ExpandProperty Value | Sort-Object -Unique | ForEach-Object {
                        $browserRows.Add([PSCustomObject]@{ Browser=$b.Name; Profile='Default'; URL=$_; Method='BinaryExtract' })
                    }
                Remove-Item $dbTmp -Force -EA SilentlyContinue
            }
        }

        # IE / Legacy Edge typed URLs (registry)
        $typedUrls = Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Internet Explorer\TypedURLs' -EA SilentlyContinue
        if ($typedUrls) {
            $typedUrls.PSObject.Properties | Where-Object { $_.Name -match '^url' } | ForEach-Object {
                $browserRows.Add([PSCustomObject]@{ Browser='IE/EdgeLegacy'; Profile='TypedURLs'; URL=$_.Value; Method='Registry' })
            }
        }

        Remove-Item $tmpDir -Recurse -Force -EA SilentlyContinue

        Save-Module '19_BrowserHistory' $browserRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='BrowserHistory'; Status='OK'; Records=$browserRows.Count })
    }
    catch { Write-Log "  [!]  Browser History ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='BrowserHistory'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 13 — Application Crash History (Event Log + WER)
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [13/25]  Application Crash History  ($DaysBack-day window)"
    try {
        $crashRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Application event log — crashes, hangs, .NET errors
        $appEvts = Get-WinEvent -FilterHashtable @{
            LogName='Application'; Id=@(1000,1001,1002,1026); StartTime=$cutoff
        } -EA SilentlyContinue

        foreach ($evt in $appEvts) {
            $xml  = [xml]$evt.ToXml()
            $data = @{}; $i = 0
            $xml.Event.EventData.Data | ForEach-Object {
                if ($_.Name) { $data[$_.Name] = $_.'#text' }
                else { $data["p$i"] = $_.'#text'; $i++ }
            }
            $crashRows.Add([PSCustomObject]@{
                TimeCreated   = $evt.TimeCreated
                EventId       = $evt.Id
                CrashType     = switch($evt.Id){ 1000{'App Crash'} 1001{'WER Report'} 1002{'App Hang'} 1026{'.NET Error'} default{'App Error'} }
                Application   = $(if ($data['Application'])        { $data['Application'] }        elseif ($data['p0']) { $data['p0'] } else { $evt.ProviderName })
                AppVersion    = $(if ($data['ApplicationVersion']) { $data['ApplicationVersion'] } else { $data['p1'] })
                FaultModule   = $(if ($data['FaultModuleName'])   { $data['FaultModuleName'] }   else { $data['p3'] })
                ExceptionCode = $(if ($data['ExceptionCode'])     { $data['ExceptionCode'] }     else { $data['p6'] })
                AppPath       = $(if ($data['ApplicationPath'])   { $data['ApplicationPath'] }   else { $data['p4'] })
            })
        }

        # Windows Error Reporting local archives
        @(
            "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive",
            "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue",
            "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
        ) | Where-Object { Test-Path $_ } | ForEach-Object {
            Get-ChildItem $_ -Directory -EA SilentlyContinue |
                Where-Object { $_.CreationTime -ge $cutoff } |
                ForEach-Object {
                    $werFolder = $_
                    $appName   = ($werFolder.Name -split '_' | Select-Object -First 1)
                    $crashRows.Add([PSCustomObject]@{
                        TimeCreated   = $werFolder.CreationTime
                        EventId       = 'WER'
                        CrashType     = 'WER Archive Report'
                        Application   = $appName
                        AppVersion    = ''
                        FaultModule   = ''
                        ExceptionCode = ''
                        AppPath       = $werFolder.FullName
                    })
                }
        }

        Save-Module '20_AppCrashHistory' ($crashRows | Sort-Object TimeCreated -Descending) $reportDir
        $summary.Add([PSCustomObject]@{ Module='CrashHistory'; Status='OK'; Records=$crashRows.Count })
    }
    catch { Write-Log "  [!]  Crash History ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='CrashHistory'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 14 — Important Registry Settings
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [14/25]  Important Registry Settings"
    try {
        $regRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        $regRoots = @(
            'HKLM:\SOFTWARE\Microsoft',
            'HKLM:\SOFTWARE\Policies',
            'HKLM:\SYSTEM\CurrentControlSet'
        )

        foreach ($regRoot in $regRoots) {
            if (-not (Test-Path $regRoot)) { continue }
            # Export top-level values from the root key itself
            try {
                $props = Get-ItemProperty -Path $regRoot -EA Stop
                $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                    $regRows.Add([PSCustomObject]@{
                        HivePath = $regRoot; SubKey = '(root)'; ValueName = $_.Name; ValueData = "$($_.Value)"; ValueType = $_.TypeNameOfValue
                    })
                }
            } catch {}
            # Export values from immediate child keys (depth 1) to keep output manageable
            Get-ChildItem -Path $regRoot -EA SilentlyContinue | ForEach-Object {
                $subKeyPath = $_.PSPath
                $subKeyName = $_.PSChildName
                try {
                    $props = Get-ItemProperty -Path $subKeyPath -EA Stop
                    $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                        $regRows.Add([PSCustomObject]@{
                            HivePath = $regRoot; SubKey = $subKeyName; ValueName = $_.Name; ValueData = "$($_.Value)"; ValueType = $_.TypeNameOfValue
                        })
                    }
                } catch {}
            }
        }

        Save-Module '21_RegistrySettings' $regRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='RegistrySettings'; Status='OK'; Records=$regRows.Count })
    }
    catch { Write-Log "  [!]  Registry Settings ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='RegistrySettings'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 15 — User Profile File Listing
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [15/25]  User Profile File Listing"
    try {
        $profileRoot = $env:USERPROFILE
        $profileFiles = Get-ChildItem -Path $profileRoot -Recurse -File -EA SilentlyContinue |
            Select-Object @{N='RelativePath';E={ $_.FullName.Substring($profileRoot.Length) }},
                Name, Extension,
                @{N='SizeKB';E={ [math]::Round($_.Length/1KB, 2) }},
                CreationTime, LastWriteTime, LastAccessTime, Attributes

        Save-Module '22_UserProfileFiles' $profileFiles $reportDir
        $summary.Add([PSCustomObject]@{ Module='UserProfileFiles'; Status='OK'; Records=($profileFiles | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  User Profile Files ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='UserProfileFiles'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 16 — Firewall Rules
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [16/25]  Firewall Rules"
    try {
        # Bulk-fetch all filters once (3 CIM calls total instead of 3 per rule)
        $portHash = @{}; Get-NetFirewallPortFilter -All -EA SilentlyContinue | ForEach-Object { $portHash[$_.InstanceID] = $_ }
        $addrHash = @{}; Get-NetFirewallAddressFilter -All -EA SilentlyContinue | ForEach-Object { $addrHash[$_.InstanceID] = $_ }
        $appHash  = @{}; Get-NetFirewallApplicationFilter -All -EA SilentlyContinue | ForEach-Object { $appHash[$_.InstanceID] = $_ }

        $fwRules = Get-NetFirewallRule -EA SilentlyContinue | ForEach-Object {
            $id = $_.InstanceID
            $pf = $portHash[$id]
            $af = $addrHash[$id]
            $ap = $appHash[$id]
            [PSCustomObject]@{
                DisplayName   = $_.DisplayName
                Name          = $_.Name
                Enabled       = $_.Enabled
                Direction     = $_.Direction
                Action        = $_.Action
                Profile       = $_.Profile
                Protocol      = $pf.Protocol
                LocalPort     = $pf.LocalPort
                RemotePort    = $pf.RemotePort
                LocalAddress  = $af.LocalAddress
                RemoteAddress = $af.RemoteAddress
                Program       = $ap.Program
                Group         = $_.Group
                Description   = $_.Description
            }
        } | Sort-Object Enabled, Direction, DisplayName

        Save-Module '23_FirewallRules' $fwRules $reportDir
        $summary.Add([PSCustomObject]@{ Module='FirewallRules'; Status='OK'; Records=($fwRules | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Firewall Rules ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='FirewallRules'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 17 — Group Policy (RSoP)
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [17/25]  Group Policy Export"
    if (-not $script:IsAdmin) {
        Write-Log "  --  SKIPPED (Group Policy export requires admin privileges)"
        $summary.Add([PSCustomObject]@{ Module='GroupPolicy'; Status='Skipped (not elevated)'; Records=0 })
    } else {
    try {
        # Generate HTML report via gpresult
        $gpHtml = Join-Path $reportDir '24_GroupPolicy.html'
        $gpProc = Start-Process -FilePath 'gpresult.exe' -ArgumentList "/H `"$gpHtml`" /F" -WindowStyle Hidden -Wait -PassThru -EA Stop
        if (Test-Path $gpHtml) {
            Write-Log "  [OK]  24_GroupPolicy.html written"
        } else {
            Write-Log "  --  gpresult HTML not generated (exit code: $($gpProc.ExitCode))"
        }

        # Also export a CSV-friendly summary from RSOP WMI
        $gpRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Applied GPOs
        Get-CimInstance -Namespace 'ROOT\RSOP\Computer' -ClassName 'RSOP_GPO' -EA SilentlyContinue | ForEach-Object {
            $gpRows.Add([PSCustomObject]@{
                Scope    = 'Computer'
                GPOName  = $_.Name
                GUID     = $_.GUIDName
                ID       = $_.ID
                Enabled  = $_.Enabled
                AccessDenied = $_.AccessDenied
                Version  = $_.Version
            })
        }
        Get-CimInstance -Namespace 'ROOT\RSOP\User' -ClassName 'RSOP_GPO' -EA SilentlyContinue | ForEach-Object {
            $gpRows.Add([PSCustomObject]@{
                Scope    = 'User'
                GPOName  = $_.Name
                GUID     = $_.GUIDName
                ID       = $_.ID
                Enabled  = $_.Enabled
                AccessDenied = $_.AccessDenied
                Version  = $_.Version
            })
        }

        Save-Module '24_GroupPolicySummary' $gpRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='GroupPolicy'; Status='OK'; Records=$gpRows.Count })
    }
    catch { Write-Log "  [!]  Group Policy ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='GroupPolicy'; Status="Error: $_"; Records=0 }) }
    }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 18 — Windows Update Logs
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [18/25]  Windows Update Logs"
    try {
        $wuRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Method 1: Query Windows Update COM object for update history
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $histCount = $searcher.GetTotalHistoryCount()
        if ($histCount -gt 0) {
            $searcher.QueryHistory(0, $histCount) | ForEach-Object {
                $wuRows.Add([PSCustomObject]@{
                    Source       = 'WUHistory'
                    Date         = $_.Date
                    Title        = $_.Title
                    Description  = $_.Description
                    ResultCode   = switch([int]$_.ResultCode){ 0{'NotStarted'} 1{'InProgress'} 2{'Succeeded'} 3{'SucceededWithErrors'} 4{'Failed'} 5{'Aborted'} default{$_.ResultCode} }
                    HResult      = $(if ($_.HResult) { '0x{0:X8}' -f $_.HResult } else { '' })
                    UpdateID     = $_.UpdateIdentity.UpdateID
                    SupportUrl   = $_.SupportUrl
                })
            }
        }

        # Method 2: Copy the WindowsUpdate.log if it exists (older Win10 / Server)
        $wuLog = "$env:SystemRoot\WindowsUpdate.log"
        if (Test-Path $wuLog) {
            Copy-Item $wuLog (Join-Path $reportDir '25b_WindowsUpdate.log') -Force -EA SilentlyContinue
            Write-Log "  [OK]  25b_WindowsUpdate.log copied"
        }

        # Method 3: On Win10+, attempt Get-WindowsUpdateLog to generate a readable log (requires admin)
        if ($script:IsAdmin) {
            $wuGenLog = Join-Path $reportDir '25c_WindowsUpdateGenerated.log'
            try {
                Get-WindowsUpdateLog -LogPath $wuGenLog -EA Stop | Out-Null
                if (Test-Path $wuGenLog) {
                    Write-Log "  [OK]  25c_WindowsUpdateGenerated.log written"
                }
            } catch {
                Write-Log "  --  Get-WindowsUpdateLog not available or failed"
            }
        } else {
            Write-Log "  --  Get-WindowsUpdateLog skipped (requires admin privileges)"
        }

        Save-Module '25_WindowsUpdateHistory' $wuRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='WindowsUpdateLogs'; Status='OK'; Records=$wuRows.Count })
    }
    catch { Write-Log "  [!]  Windows Update Logs ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='WindowsUpdateLogs'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 19 — Application Event Log Errors
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [19/25]  Application Event Log Errors  ($DaysBack-day window)"
    try {
        # Level 1 = Critical, Level 2 = Error
        $appErrors = Get-WinEvent -FilterHashtable @{
            LogName='Application'; Level=@(1,2); StartTime=$cutoff
        } -EA SilentlyContinue | ForEach-Object {
            [PSCustomObject]@{
                TimeCreated      = $_.TimeCreated
                EventId          = $_.Id
                Level            = $_.LevelDisplayName
                ProviderName     = $_.ProviderName
                MachineName      = $_.MachineName
                UserId           = $(if ($_.UserId) { $_.UserId.Value } else { '' })
                Message          = $(if ($_.Message) { $m = $_.Message -replace '\s+',' '; $m.Substring(0, [Math]::Min(500, $m.Length)) } else { '' })
            }
        } | Sort-Object TimeCreated -Descending

        Save-Module '26_AppEventLogErrors' $appErrors $reportDir
        $summary.Add([PSCustomObject]@{ Module='AppEventLogErrors'; Status='OK'; Records=($appErrors | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  App Event Log Errors ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='AppEventLogErrors'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 20 — System Energy Report
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [20/25]  System Energy Report"
    if (-not $script:IsAdmin) {
        Write-Log "  --  SKIPPED (Energy report requires admin privileges)"
        $summary.Add([PSCustomObject]@{ Module='EnergyReport'; Status='Skipped (not elevated)'; Records=0 })
    } else {
    try {
        $energyXml  = Join-Path $env:TEMP 'energy-report.xml'

        # powercfg /energy outputs an XML report to %TEMP% by default; /OUTPUT sets destination
        # Run a short 10-second trace to keep it quick
        $pcfgProc = Start-Process -FilePath 'powercfg.exe' -ArgumentList "/energy /duration 10 /output `"$energyXml`"" -WindowStyle Hidden -Wait -PassThru -EA Stop

        if (Test-Path $energyXml) {
            # Copy raw XML to report folder
            Copy-Item $energyXml (Join-Path $reportDir '27_EnergyReport.xml') -Force -EA SilentlyContinue

            # Parse key findings into CSV
            $energyRows = [System.Collections.Generic.List[PSCustomObject]]::new()
            try {
                [xml]$eXml = Get-Content $energyXml -Raw -EA Stop
                $eXml.EnergyReport.Warnings.Warning | ForEach-Object {
                    $energyRows.Add([PSCustomObject]@{
                        Severity    = 'Warning'
                        Category    = $_.Category
                        Name        = $_.Name
                        Description = $_.Description
                    })
                }
                $eXml.EnergyReport.Errors.Error | ForEach-Object {
                    $energyRows.Add([PSCustomObject]@{
                        Severity    = 'Error'
                        Category    = $_.Category
                        Name        = $_.Name
                        Description = $_.Description
                    })
                }
                $eXml.EnergyReport.Informational.Info | ForEach-Object {
                    $energyRows.Add([PSCustomObject]@{
                        Severity    = 'Info'
                        Category    = $_.Category
                        Name        = $_.Name
                        Description = $_.Description
                    })
                }
            } catch {
                Write-Log "  --  Energy XML parse issue (raw XML still saved)"
            }

            Save-Module '27_EnergyReport' $energyRows.ToArray() $reportDir
            $summary.Add([PSCustomObject]@{ Module='EnergyReport'; Status='OK'; Records=$energyRows.Count })

            Remove-Item $energyXml -Force -EA SilentlyContinue
        } else {
            Write-Log "  --  powercfg /energy did not produce output (exit code: $($pcfgProc.ExitCode))"
            $summary.Add([PSCustomObject]@{ Module='EnergyReport'; Status='No output'; Records=0 })
        }
    }
    catch { Write-Log "  [!]  Energy Report ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='EnergyReport'; Status="Error: $_"; Records=0 }) }
    }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 21 — Proxy Configuration
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [21/25]  Proxy Configuration"
    try {
        $proxyRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        # User-level proxy (Internet Settings)
        $inetSettings = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -EA SilentlyContinue
        if ($inetSettings) {
            $proxyRows.Add([PSCustomObject]@{
                Scope         = 'User'
                Source        = 'InternetSettings'
                ProxyEnabled  = $inetSettings.ProxyEnable
                ProxyServer   = $inetSettings.ProxyServer
                ProxyOverride = $inetSettings.ProxyOverride
                AutoConfigURL = $inetSettings.AutoConfigURL
            })
        }

        # System-level proxy (WinHTTP)
        try {
            $winhttp = & netsh winhttp show proxy 2>&1
            $proxyRows.Add([PSCustomObject]@{
                Scope         = 'System'
                Source        = 'WinHTTP'
                ProxyEnabled  = $(if ($winhttp -match 'Direct access') { 'No' } else { 'Yes' })
                ProxyServer   = $(if ($winhttp -match 'Proxy Server.*:\s*(.+)') { $Matches[1].Trim() } else { '' })
                ProxyOverride = $(if ($winhttp -match 'Bypass List.*:\s*(.+)')  { $Matches[1].Trim() } else { '' })
                AutoConfigURL = ''
            })
        } catch {}

        # Environment variable proxies
        foreach ($envVar in @('HTTP_PROXY','HTTPS_PROXY','NO_PROXY','ALL_PROXY')) {
            $val = [Environment]::GetEnvironmentVariable($envVar, 'User')
            $sysVal = [Environment]::GetEnvironmentVariable($envVar, 'Machine')
            if ($val) {
                $proxyRows.Add([PSCustomObject]@{
                    Scope='User'; Source="EnvVar:$envVar"; ProxyEnabled='Yes'
                    ProxyServer=$val; ProxyOverride=''; AutoConfigURL=''
                })
            }
            if ($sysVal) {
                $proxyRows.Add([PSCustomObject]@{
                    Scope='System'; Source="EnvVar:$envVar"; ProxyEnabled='Yes'
                    ProxyServer=$sysVal; ProxyOverride=''; AutoConfigURL=''
                })
            }
        }

        # Group Policy proxy settings
        $gpProxy = Get-ItemProperty 'HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -EA SilentlyContinue
        if ($gpProxy -and ($gpProxy.ProxyEnable -or $gpProxy.ProxyServer)) {
            $proxyRows.Add([PSCustomObject]@{
                Scope         = 'GroupPolicy'
                Source        = 'GPO_InternetSettings'
                ProxyEnabled  = $gpProxy.ProxyEnable
                ProxyServer   = $gpProxy.ProxyServer
                ProxyOverride = $gpProxy.ProxyOverride
                AutoConfigURL = $gpProxy.AutoConfigURL
            })
        }

        Save-Module '28_ProxyConfiguration' $proxyRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='ProxyConfig'; Status='OK'; Records=$proxyRows.Count })
    }
    catch { Write-Log "  [!]  Proxy Config ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='ProxyConfig'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 22 — BitLocker Configuration & Status
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [22/25]  BitLocker Configuration & Status"
    if (-not $script:IsAdmin) {
        Write-Log "  --  SKIPPED (BitLocker status requires admin privileges)"
        $summary.Add([PSCustomObject]@{ Module='BitLocker'; Status='Skipped (not elevated)'; Records=0 })
    } else {
    try {
        $blRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        $blVolumes = Get-BitLockerVolume -EA Stop
        foreach ($vol in $blVolumes) {
            $blRows.Add([PSCustomObject]@{
                MountPoint       = $vol.MountPoint
                VolumeType       = $vol.VolumeType
                ProtectionStatus = $vol.ProtectionStatus
                LockStatus       = $vol.LockStatus
                EncryptionMethod = $vol.EncryptionMethod
                EncryptionPercent = $vol.EncryptionPercentage
                VolumeStatus     = $vol.VolumeStatus
                KeyProtectors    = ($vol.KeyProtector | ForEach-Object { "$($_.KeyProtectorType)($($_.KeyProtectorId))" }) -join '; '
                AutoUnlockEnabled = $vol.AutoUnlockEnabled
                AutoUnlockKeyStored = $vol.AutoUnlockKeyStored
                CapacityGB       = $([math]::Round($vol.CapacityGB, 2))
            })
        }

        Save-Module '29_BitLockerStatus' $blRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='BitLocker'; Status='OK'; Records=$blRows.Count })
    }
    catch {
        if ($_.Exception.Message -match 'not recognized|not found|not loaded') {
            Write-Log "  --  BitLocker module not available on this edition of Windows"
            $summary.Add([PSCustomObject]@{ Module='BitLocker'; Status='Not available'; Records=0 })
        } else {
            Write-Log "  [!]  BitLocker ERROR: $_"
            $summary.Add([PSCustomObject]@{ Module='BitLocker'; Status="Error: $_"; Records=0 })
        }
    }
    }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 23 — Disk Health / SMART Data
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [23/25]  Disk Health / SMART Data"
    try {
        $diskRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        $physDisks = Get-PhysicalDisk -EA SilentlyContinue
        foreach ($pd in $physDisks) {
            $rel = $pd | Get-StorageReliabilityCounter -EA SilentlyContinue
            $diskRows.Add([PSCustomObject]@{
                DeviceId         = $pd.DeviceId
                FriendlyName     = $pd.FriendlyName
                MediaType        = $pd.MediaType
                BusType          = $pd.BusType
                HealthStatus     = $pd.HealthStatus
                OperationalStatus = $pd.OperationalStatus
                SizeGB           = $([math]::Round($pd.Size / 1GB, 2))
                FirmwareVersion  = $pd.FirmwareVersion
                SerialNumber     = $pd.SerialNumber
                Temperature      = $rel.Temperature
                ReadErrorsTotal  = $rel.ReadErrorsTotal
                WriteErrorsTotal = $rel.WriteErrorsTotal
                PowerOnHours     = $rel.PowerOnHours
                Wear             = $rel.Wear
            })
        }

        # Fallback: also grab WMI disk info for additional SMART context
        Get-CimInstance -Namespace 'ROOT\WMI' -ClassName 'MSStorageDriver_FailurePredictStatus' -EA SilentlyContinue | ForEach-Object {
            $diskRows.Add([PSCustomObject]@{
                DeviceId='WMI'; FriendlyName=$_.InstanceName
                MediaType=''; BusType=''; HealthStatus=''
                OperationalStatus=$(if ($_.PredictFailure) { 'FAILURE PREDICTED' } else { 'OK' })
                SizeGB=''; FirmwareVersion=''; SerialNumber=''
                Temperature=''; ReadErrorsTotal=''; WriteErrorsTotal=''
                PowerOnHours=''; Wear=''
            })
        }

        Save-Module '30_DiskHealthSMART' $diskRows.ToArray() $reportDir
        $summary.Add([PSCustomObject]@{ Module='DiskHealth'; Status='OK'; Records=$diskRows.Count })
    }
    catch { Write-Log "  [!]  Disk Health ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='DiskHealth'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 24 — Installed Driver Versions
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [24/25]  Installed Driver Versions"
    try {
        $drivers = Get-CimInstance Win32_PnPSignedDriver -EA SilentlyContinue |
            Where-Object { $_.DeviceName } |
            Select-Object DeviceName, DeviceClass, DriverVersion, DriverDate,
                DriverProviderName, Manufacturer, InfName, IsSigned, Signer,
                @{N='HardwareID';E={ $_.HardwareID | Select-Object -First 1 }} |
            Sort-Object DeviceClass, DeviceName

        Save-Module '31_InstalledDrivers' $drivers $reportDir
        $summary.Add([PSCustomObject]@{ Module='InstalledDrivers'; Status='OK'; Records=($drivers | Measure-Object).Count })
    }
    catch { Write-Log "  [!]  Installed Drivers ERROR: $_"; $summary.Add([PSCustomObject]@{ Module='InstalledDrivers'; Status="Error: $_"; Records=0 }) }
    Step-Progress

    # ────────────────────────────────────────────────────────────────────────
    #  MODULE 25 — Audit Summary CSV
    # ────────────────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "> [25/25]  Finalising Audit Summary"
    try {
        $summary.Add([PSCustomObject]@{ Module='-- METADATA --'; Status='Computer'; Records=$hostname })
        $summary.Add([PSCustomObject]@{ Module='RunBy';      Status=$env:USERNAME;                 Records='' })
        $summary.Add([PSCustomObject]@{ Module='AuditTime';  Status=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Records='' })
        $summary.Add([PSCustomObject]@{ Module='DaysBack';   Status=$DaysBack;                     Records='' })
        $summary.Add([PSCustomObject]@{ Module='ReportPath'; Status=$reportDir;                    Records='' })
        $summary | Export-Csv -Path "$reportDir\00_AuditSummary.csv" -NoTypeInformation -Encoding UTF8
        Write-Log "  [OK]  00_AuditSummary.csv written"
        Step-Progress
    }
    catch { Write-Log "  [!]  Summary ERROR: $_" }

    # ── Complete ──────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "=============================================="
    Write-Log "  [OK]  EXPORT COMPLETE"
    Write-Log "  $reportDir"
    Write-Log "=============================================="

    $script:SH.Dispatcher.Invoke([action]{
        $script:SH.Status.Text  = "Export complete -- report saved."
        $script:SH.RunBtn.Content    = "$([char]0x25B6)  Run Export"
        $script:SH.RunBtn.IsEnabled  = $true
        $script:SH.OpenBtn.IsEnabled = $true
        $script:SH.OpenBtn.Visibility = [System.Windows.Visibility]::Visible
        $script:SH.PBar.Value = 100
    }, 'Normal')
}
#endregion

#region ── Run Export Button ───────────────────────────────────────────────────
$BtnRun.Add_Click({
    $outPath  = $TxtOutput.Text
    $daysText = $TxtDays.Text

    if (-not $outPath -or -not (Test-Path $outPath)) {
        [System.Windows.MessageBox]::Show(
            "Please select a valid output folder.", "SystemLogExporter",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }
    if ($daysText -notmatch '^\d+$' -or [int]$daysText -lt 1 -or [int]$daysText -gt 3650) {
        [System.Windows.MessageBox]::Show(
            "Look-back days must be a positive integer (1-3650).", "SystemLogExporter",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $days = [int]$daysText

    # Reset UI
    $TxtLog.Text                   = ''
    $PBar.Value                    = 0
    $BtnRun.IsEnabled              = $false
    $BtnRun.Content                = "Running..."
    $BtnOpen.Visibility            = [System.Windows.Visibility]::Collapsed
    $BtnOpen.IsEnabled             = $false
    $TxtStatus.Text                = "Export running..."
    $syncHash.ReportPath           = ''

    # Launch in background runspace
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('syncHash', $syncHash)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $ps.AddScript($auditScript)      | Out-Null
    $ps.AddArgument($outPath)        | Out-Null
    $ps.AddArgument($days)           | Out-Null
    $ps.AddArgument($syncHash)       | Out-Null
    $asyncResult = $ps.BeginInvoke()

    # Register a callback to dispose resources when the runspace completes
    Register-ObjectEvent -InputObject $ps -EventName InvocationStateChanged -Action {
        if ($Sender.InvocationStateInfo.State -in 'Completed','Failed','Stopped') {
            try { $Sender.EndInvoke($asyncResult) } catch {}
            $Sender.Dispose()
            $Sender.Runspace.Dispose()
            $Event.SourceObject = $null
            Unregister-Event -SubscriptionId $EventSubscriber.SubscriptionId
        }
    } | Out-Null
})
#endregion

# ── Defaults ─────────────────────────────────────────────────────────────────
$TxtOutput.Text   = [Environment]::GetFolderPath('Desktop')
$BtnRun.IsEnabled = $true
$TxtStatus.Text   = "Output set to Desktop -- click Run Export to begin."

# ── Show ─────────────────────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null
