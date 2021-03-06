 #Disk Space Utility for assisting users to know what files\folder can be cleaned up to gain space.
 #04-16  
 #Jon Carracino

 
 #Global var:
 $ProfileDir = $env:USERPROFILE
$dirRoot = $ProfileDir # Change this as desired 
$filter = '*' # Change this as desired, e.g. *.log or *.txt
$computername = $env:COMPUTERNAME
$driveletter = $env:HOMEDRIVE

 # Disk space info function:
 Function Get-DiskInfo {
$computername =$env:COMPUTERNAME
Get-WMIObject Win32_logicaldisk -ComputerName $computername | Select-Object @{Name='ComputerName';Ex={$computername}},`
                                                                    @{Name=‘Drive Letter‘;Expression={$_.DeviceID}},`
                                                                    @{Name=‘Drive Label’;Expression={$_.VolumeName}},`
                                                                    @{Name=‘Size(MB)’;Expression={[int]($_.Size / 1MB)}},`
                                                                    @{Name=‘FreeSpace%’;Expression={[math]::Round($_.FreeSpace / $_.Size,2)*100}}
                                                                 }
     
 # FolderSize Function
function GetFolderSize($path){ 
    $total = (Get-ChildItem $path -ErrorAction SilentlyContinue -filter $filter | Measure-Object -Property length -Sum -ErrorAction SilentlyContinue).Sum 
    if (-not($total)) { $total = 0 } 
    $total 
    } # end function GetFolderSize 

###### Start of Form Building:	
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="Disk Space Info" 
    $MyForm.Size = New-Object System.Drawing.Size(500,300)
	#$MyForm.AutoSize = $True
	#$MyForm.AutoSizeMode = "GrowAndShrink"
	$MyForm.StartPosition = "CenterScreen"
	
	#Ensure exits cleanly with correct ExitCode:
	#$MyForm.add_FormClosing([System.Windows.Forms.FormClosingEventHandler]{ 
	#	$script:ExitCode = 0 #Set the exit code for the Packager
	#}) 
	
	#get working dir
	$invocation = (Get-Variable MyInvocation).Value
	$directorypath = Split-Path $invocation.MyCommand.Path
	#$settingspath = $directorypath + '\CapOne.ico'
	$MyForm.Icon = New-Object System.Drawing.Icon("$directorypath\CapOne.ico")
     
 	#Add Top Label:
        $mLabel1 = New-Object System.Windows.Forms.Label 
                $mLabel1.Text="What's taking up my space?"
				$mLabel1.AutoSize = $True
                $mLabel1.Top="35" 
                $mLabel1.Left="173" 
                $mLabel1.Anchor="Left,Top" 		  
        $MyForm.Controls.Add($mLabel1) 
		
	#Add Info Label:	
         $mLabel2 = New-Object System.Windows.Forms.Label 
                $mLabel2.Text="In order to maintain proper performance, you should have at least 15% free space."
				$mLabel2.AutoSize = $True
                $mLabel2.Top="10" 
                $mLabel2.Left="50" 
                $mLabel2.Anchor="Left,Top" 		  
        $MyForm.Controls.Add($mLabel2) 
 
        $mDataGrid1 = New-Object System.Windows.Forms.DataGrid 
                $mDataGrid1.Text="DataGrid1" 
                $mDataGrid1.Top="70" 
                $mDataGrid1.Left="15" 
                $mDataGrid1.Anchor="Left,Top" 
        $mDataGrid1.Size = New-Object System.Drawing.Size(450,110) 
        $MyForm.Controls.Add($mDataGrid1) 
         
         #Create Button1 : FileInfo
        $mButton1 = New-Object System.Windows.Forms.Button 
                $mButton1.Text="Show Me Large Files" 
                $mButton1.Top="200" 
                $mButton1.Left="67" 
                $mButton1.Anchor="Left,Top" 
        $mButton1.Size = New-Object System.Drawing.Size(150,23) 
        $MyForm.Controls.Add($mButton1) 
		
		 #Create Button2 : FolderInfo
        $mButton2 = New-Object System.Windows.Forms.Button 
                $mButton2.Text="Show Me Large Folders" 
                $mButton2.Top="200" 
                $mButton2.Left="267" 
                $mButton2.Anchor="Left,Top" 
        $mButton2.Size = New-Object System.Drawing.Size(150,23) 
        $MyForm.Controls.Add($mButton2) 
		
		$diskInfo = [system.collections.arraylist](Get-WMIObject Win32_logicaldisk -ComputerName $computername | Where-Object {$_.DeviceID -eq "B:" -or $_.DeviceID -eq "C:" -or $_.DeviceID -eq $driveletter} | Select-Object @{Name='ComputerName';Ex={$computername}},`
                                                                    @{Name=‘Drive Letter‘;Expression={$_.DeviceID}},`
                                                                    @{Name=‘Drive Label’;Expression={$_.VolumeName}},`
                                                                    @{Name=‘Size(MB)’;Expression={[int]($_.Size / 1MB)}},`
                                                                    @{Name=‘FreeSpace%’;Expression={[math]::Round($_.FreeSpace / $_.Size,2)*100}})
		#$diskInfo2 is what gets current drive space details
		#Then puts data into gridformat for GUI.
		$mDataGrid1.DataSource = $diskInfo
		#Write-Host $diskInfo
					
		
##############Button1 - Show File Info:
		$mButton1.add_Click({
               #Export csv to desktop
			   #Get-ChildItem -Path "$ProfileDir" * -Recurse | where-object {$_.Length -gt 1000000} | Select-object -Property Name, Length, Directory, CreationTime, LastAccessTime, @{Name='Size';Expression={$_.Length / 1MB}} | Sort-Object -Property Length -Descending | Export-Csv -Path "$ProfileDir\Desktop\FileInfo.csv" -NoTypeInformation -Force                                                  
				Get-ChildItem -Path "$ProfileDir", "$ProfileDir\AppData\Local\Microsoft" * -Recurse | where-object {$_.Length -gt 10000000} | Select-object -Property Name, @{Name='Size(MB)';Expression={$_.Length / 1MB}}, Directory, CreationTime, LastAccessTime | Sort-Object -Property 'Size(MB)' -Descending | Out-GridView -Title "List of large files inside profile sorted by size"                                               
		})
		
##########Button2 - Show folder info:
		$mButton2.add_Click({
		# Entry point into script using global var values at start of script....
		$results = @() 
		$dirs = Get-ChildItem $dirRoot -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.psIsContainer} 
 
		foreach ($dir in $dirs) { 
    
   		 $childFiles = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue -filter $filter| Where-Object{ -not($_.psIsContainer)}) 
    	if ($childFiles) { $filecount = ($childFiles.count)} 
    		else                     { $filecount = 0                  } 
 
    	$childDirs = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue | Where-Object{ $_.psIsContainer}) 
    	if ($childDirs ){ $dircount = ($childDirs.count)} 
    		else                    { $dircount = 0                 } 
     
    	$result = New-Object psobject -Property @{Folder = (Split-Path $dir.pspath -NoQualifier) 
                                              TotalSize = (GetFolderSize($dir.pspath)) 
                                              FileCount = $filecount; SubDirs = $dircount} 
    	$results += $result 
	    } # end foreach 
 			#####Display results as a grid in seperate window...
			$results | Select-Object Folder, TotalSize , FileCount, SubDirs |where-object {$_.TotalSize -gt 10000000} | Sort-Object TotalSize -Descending | Out-GridView -Title "List of large folders inside profile sorted by size" 
		})
###############  End Button2 - Folder Size

        $MyForm.ShowDialog() | Out-Null
		
		#$script:ExitCode = 0

