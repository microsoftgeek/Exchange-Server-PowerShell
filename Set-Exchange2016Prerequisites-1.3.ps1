<#  
.SYNOPSIS
   	Configures the necessary prerequisites to install Exchange 2016 on a Windows Server 2012 (R2) server.

.DESCRIPTION  
    Installs all required Windows Server 2012 (R2) components, downloading latest Update Rollup, etc.

.NOTES  
    Version      		: 1.3
    Change Log			: 1.3 - Added SSL Security enhancements (optional)
				: 1.2 - Added High Performance Power Plan change, cleaned up menu
				: 1.1 - Added NIC Power Management
				: 1.0 First iteration
    Wish list			: better comment based help
				: event log logging
    Rights Required		: Local admin on server
    Sched Task Req'd		: No
    Exchange Version		: 2016
    Author       		: Just A UC Guy [Damian Scoles]
    Dedicated Blog		: http://justaucguy.wordpress.com
    Disclaimer   		: You are on your own.  This was not written by, support by, or endorsed by Microsoft.
    Info Stolen from 		: Anderson Patricio, Bhargav Shukla and Pat Richard [Exchange 2010 script]
    				: http://msmvps.com/blogs/andersonpatricio/archive/2009/11/13/installing-exchange-server-2010-pre-requisites-on-windows-server-2008-r2.aspx
				: http://www.bhargavs.com/index.php/powershell/2009/11/script-to-install-exchange-2010-pre-requisites-for-windows-server-2008-r2/
				: SQL Soldier - http://www.sqlsoldier.com/wp/sqlserver/enabling-high-performance-power-plan-via-powershell
.LINK  
[TBD]

.EXAMPLE
	.\Set-Exchange2016Prerequisites-1.2.ps1

.INPUTS
	None. You cannot pipe objects to this script.
#>
#Requires -Version 2.0
param(
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[string] $strFilenameTranscript = $MyInvocation.MyCommand.Name + " " + (hostname)+ " {0:yyyy-MM-dd hh-mmtt}.log" -f (Get-Date),
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true, Mandatory=$false)] 
	[string] $TargetFolder = "c:\Install",
	# [string] $TargetFolder = $Env:Temp
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $WasInstalled = $false,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $RebootRequired = $false,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[string] $opt = "None",
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
)

Start-Transcript -path .\$strFilenameTranscript | Out-Null
$error.clear()
# Detect correct OS here and exit if no match (we intentionally truncate the last character to account for service packs)

# ******************************************************
# *    This section is for the Windows 2012 (R2) OS    *
# ******************************************************

if ((Get-WMIObject win32_OperatingSystem).Version -notmatch '6.2'){
	if ((Get-WMIObject win32_OperatingSystem).Version -notmatch '6.3'){
	Write-Host "`nThis script requires a version of Windows Server 2012 or 2012 R2, which this is not. Exiting...`n" -ForegroundColor Red
	Exit
	}
}
Clear-Host
Pushd
# determine if BitsTransfer is already installed
if ((Get-Module BitsTransfer).installed -eq $true){
	[bool] $WasInstalled = $true
}else{
	[bool] $WasInstalled = $false
}
[string] $menu = @'

	******************************************************************
	Exchange Server 2016 [On Windows 2012 (R2)] - Features script
	******************************************************************
	
	Please select an option from the list below:

    	1) Install Mailbox prerequisites - Part 1 (Includes Option 30/31 below)
    	2) Install Mailbox prerequisites - Part 2
    	3) Install Edge Transport Server prerequisites

    	10) Launch Windows Update
    	11) Check Prerequisites for Mailbox role
    	12) Check Prerequisites for Edge role

    	20) Install - One-Off - .NET 4.5.2 [MBX or Edge]
    	21) Install - One-Off - Windows Features [MBX]
    	22) Install - One Off - Unified Communications Managed API 4.0

	30) Set Power Plan to High Performance (Recommended by MS)
	31) Disable Power Management for NICs.
	32) Disable SSL 3.0 Support     ** NEW **
	33) Disable RC4 Support     ** NEW **
    	
	98) Restart the Server
	99) Exit

Select an option.. [1-99]?
'@

function highperformance {
	Try {
        	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	        $CurrPlan = $(powercfg -getactivescheme).split()[3]
        	if ($CurrPlan -ne $HighPerf) {
			powercfg -setactive $HighPerf
			if ($CurrPlan -eq $HighPerf) {
				write-host " ";write-host "The power plan now is set to " -nonewline;write-host "High Performance." -foregroundcolor green;write-host " "
			}
		} else {
			if ($CurrPlan -eq $HighPerf) {
				write-host " ";write-host "The power plan is already set to " -nonewline;write-host "High Performance." -foregroundcolor green;write-host " "
			}
		}
	    } Catch {
        	Write-Warning -Message "Unable to set power plan to high performance"
	    }

}

# Function - .NET 4.5.2
function Install-DotNET452{
    # .NET 4.5.2
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if ($val.Release -lt "379893") {
    		GetIt "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
	    	Set-Location $targetfolder
    		[string]$expression = ".\NDP452-KB2901907-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $targetfolder\DotNET452.log"
	    	Write-Host "File: NDP452-KB2901907-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
    		Invoke-Expression $expression
    		Start-Sleep -Seconds 20
    		Write-Host "`n.NET 4.5.2 is now installed" -Foregroundcolor Green
	} else {
    		Write-Host "`n.NET 4.5.2 already installed" -Foregroundcolor Green
    }
} # end Install-DotNET452

# Mailbox Role - Windows Feature requirements
function check-MBXprereq {
    # .NET 4.5.2
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if($val.Release -lt "379893") {
		write-host ".NET 4.5.2 is " -nonewline 
		write-host "not installed!" -ForegroundColor red
	}
	else {
		write-host ".NET 4.5.2 is " -nonewline
		write-host "installed." -ForegroundColor green
	}

# Windows Feature Check
	$values = @("AS-HTTP-Activation","Desktop-Experience","NET-Framework-45-Features","RPC-over-HTTP-proxy","RSAT-Clustering","RSAT-Clustering-CmdInterface","RSAT-Clustering-Mgmt","RSAT-Clustering-PowerShell","Web-Mgmt-Console","WAS-Process-Model","Web-Asp-Net45","Web-Basic-Auth","Web-Client-Auth","Web-Digest-Auth","Web-Dir-Browsing","Web-Dyn-Compression","Web-Http-Errors","Web-Http-Logging","Web-Http-Redirect","Web-Http-Tracing","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Lgcy-Mgmt-Console","Web-Metabase","Web-Mgmt-Console","Web-Mgmt-Service","Web-Net-Ext45","Web-Request-Monitor","Web-Server","Web-Stat-Compression","Web-Static-Content","Web-Windows-Auth","Web-WMI","Windows-Identity-Foundation")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "installed." -ForegroundColor green
		}else{
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "not installed!" -ForegroundColor red
		}
	}

# Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
  $val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
  if($val.DisplayVersion -ne "5.0.8308.0"){
    	if($val.DisplayVersion -ne "5.0.8132.0"){
        	if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
			write-host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
            		write-host "not installed!" -ForegroundColor red
            		write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		} else {
			write-host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
			write-host "installed." -ForegroundColor red
			write-host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red
			write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		}
	} else {
        	write-host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        	write-host "installed." -ForegroundColor red
        	write-host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red 
        	write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
        }   
   } else {
        write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        write-host "installed." -ForegroundColor green
   }
}

# Edge Transport requirement check
function check-EdgePrereq {
	
     # Windows Feature AD LightWeight Services
	$values = @("ADLDS")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "installed." -ForegroundColor green
		}else{
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "not installed!" -ForegroundColor red
		}
	}

    # .NET 4.5.2
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if($val.Release -lt "379893") {
		write-host ".NET 4.5.2 is " -nonewline 
		write-host "not installed!" -ForegroundColor red
	}
	else {
		write-host ".NET 4.5.2 is " -nonewline
		write-host "installed." -ForegroundColor green
	}
}

# Function - Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-WinUniComm4 {
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		if($val.DisplayVersion -ne "5.0.8132.0"){
			if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is not installed.  Downloading and installing now."
				Install-NewWinUniComm4
			} else {
    				Write-Host "`nAn old version of Microsoft Unified Communications Managed API 4.0 is installed."
				UnInstall-WinUniComm4
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now."
				Install-NewWinUniComm4
			}
   		} else {
   			Write-Host "`nThe Preview version of Microsoft Unified Communications Managed API 4.0 is installed."
   			UnInstall-WinUniComm4
   			Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now."
   			Install-NewWinUniComm4
		}
	} else {
		write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
		write-host "installed." -ForegroundColor green
	}
} # end Install-WinUniComm4

# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-NewWinUniComm4{
	GetIt "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
	Set-Location $targetfolder
	[string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $targetfolder\WinUniComm4.log"
	Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
	Invoke-Expression $expression
	Start-Sleep -Seconds 20
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -Foregroundcolor Green
	}
} # end Install-NewWinUniComm4

function GetIt ([string]$sourcefile)	{
	if ($HasInternetAccess){
		# check if BitsTransfer is installed
		if ((Get-Module BitsTransfer) -eq $null){
			Write-Host "BitsTransfer: Installing..." -NoNewLine
			Import-Module BitsTransfer	
			Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`binstalled!   " -ForegroundColor Green
		}
		[string] $targetfile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1) 
		if (Test-Path $targetfolder){
			Write-Host "Folder: $targetfolder exists."
		} else{
			Write-Host "Folder: $targetfolder does not exist, creating..." -NoNewline
			New-Item $targetfolder -type Directory | Out-Null
			Write-Host "`b`b`b`b`b`b`b`b`b`b`bcreated!   " -ForegroundColor Green
		}
		if (Test-Path "$targetfolder\$targetfile"){
			Write-Host "File: $targetfile exists."
		}else{	
			Write-Host "File: $targetfile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$targetfolder\$targetfile"
			Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`b`bdownloaded!   " -ForegroundColor Green
		}
	}else{
		Write-Host "Internet Access not detected. Please resolve and try again." -foregroundcolor red
	}
} # end GetIt

function UnZipIt ([string]$source, [string]$target){
	if (Test-Path "$targetfolder\$target"){
		Write-Host "File: $target exists."
	}else{
		Write-Host "File: $target doesn't exist, unzipping..." -NoNewLine
		$sh = new-object -com shell.application
		$zipfolder = $sh.namespace("$targetfolder\$source") 
		$item = $zipfolder.parsename("$target")      
		$targetfolder2 = $sh.namespace("$targetfolder")       
		Set-Location $targetfolder
		$targetfolder2.copyhere($item)
		Write-Host "`b`b`b`b`b`b`b`b`b`b`b`bunzipped!   " -ForegroundColor Green
		Remove-Item $source
	}
} # end UnZipIt

function Get-ModuleStatus { 
	param	(
		[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available
			return $false
		}
	}else {
		# module was already imported
		# Write-Host "$name module already imported"
		return $true
	}
} # end function Get-ModuleStatus

function New-FileDownload {
	param (
		[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No source file specified")] 
		[string]$SourceFile,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination folder specified")] 
    [string]$DestFolder,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination file specified")] 
    [string]$DestFile
	)
	# I should clean up the display text to be consistent with other functions
	$error.clear()
	if (!($DestFolder)){$DestFolder = $TargetFolder}
	Get-ModuleStatus -name BitsTransfer
	if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Host "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..." -NoNewline
		New-Item $DestFolder -type Directory
		Write-Host "Done! " -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host "File: $DestFile exists."
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile"
			Write-Host "Done! " -ForegroundColor Green
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # end function New-FileDownload

function CheckPowerPlan {
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -eq $HighPerf) {
		write-host " ";write-host "The power plan now is set to " -nonewline;write-host "High Performance." -foregroundcolor green;write-host " "
	}
}

function highperformance {
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -ne $HighPerf) {
		powercfg -setactive $HighPerf
		CheckPowerPlan
	} else {
		if ($CurrPlan -eq $HighPerf) {
			write-host " ";write-host "The power plan is already set to " -nonewline;write-host "High Performance." -foregroundcolor green;write-host " "
		}
	}
}


function PowerMgmt {
	$NICs = Get-WmiObject -Class Win32_NetworkAdapter|Where-Object{$_.PNPDeviceID -notlike "ROOT\*" -and $_.Manufacturer -ne "Microsoft" -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22} 
	Foreach($NIC in $NICs) {
		$NICName = $NIC.Name
		$DeviceID = $NIC.DeviceID
		If([Int32]$DeviceID -lt 10) {
			$DeviceNumber = "000"+$DeviceID 
		} Else {
			$DeviceNumber = "00"+$DeviceID
		}
		$KeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\$DeviceNumber"
  
		If(Test-Path -Path $KeyPath) {
			$PnPCapabilities = (Get-ItemProperty -Path $KeyPath).PnPCapabilities
			If($PnPCapabilities -eq 0){Set-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 | Out-Null
				write-host "Changed the NIC Power Management settings.";write-host " ";write-host "A reboot is REQUIRED!" -foregroundcolor red;write-host " "}
			If($PnPCapabilities -eq $null){Set-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 | Out-Null
				write-host "Changed the NIC Power Management settings.";write-host " ";write-host "A reboot is REQUIRED!" -foregroundcolor red;write-host " "}
			If($PnPCapabilities -eq 24) {write-host " ";write-host "Power Management has already been " -NoNewline;write-host "disabled" -ForegroundColor Green;write-host " "}
   		 } 
 	 } 
 }

function DisableRC4 {
	# Define Registry keys to look for
	$base = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\" -erroraction silentlycontinue
	$val1 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\" -erroraction silentlycontinue
	$val2 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\" -erroraction silentlycontinue
	$val3 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\" -erroraction silentlycontinue
	
	# Define Values to add
	$registryBase = "Ciphers"
	$registryPath1 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\"
	$registryPath2 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\"
	$registryPath3 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\"
	$Name = "Enabled"
	$value = "0"
	$ssl = 0
	$checkval1 = Get-Itemproperty -Path "$registrypath1" -name $name -erroraction silentlycontinue
	$checkval2 = Get-Itemproperty -Path "$registrypath2" -name $name -erroraction silentlycontinue
	$checkval3 = Get-Itemproperty -Path "$registrypath3" -name $name -erroraction silentlycontinue
    
# Formatting for output
	write-host " "

# Add missing registry keys as needed
	If ($base -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL", $true)
		$key.CreateSubKey('Ciphers')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "Ciphers" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

	If ($val1 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 128/128')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "Ciphers\RC4 128/128" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

	If ($val2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 40/128')
		$key.Close()
		New-ItemProperty -Path $registryPath2 -Name $name -Value $value
	} else {
		write-host "The " -nonewline;write-host "Ciphers\RC4 40/128" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

	If ($val3 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 56/128')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "Ciphers\RC4 56/128" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}
	
# Add the enabled value to disable RC4 Encryption
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath1 -Name $name -Value $value -force;$ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline;write-host "Enabled" -ForegroundColor green -NoNewline;write-host " exists under the RC4 128/128 Registry Key.";$ssl++
	}
	If ($checkval2.enabled -ne "0") {
		write-host $checkval2
		try {
			New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force;$ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline;write-host "Enabled" -ForegroundColor green -NoNewline;write-host " exists under the RC4 40/128 Registry Key.";$ssl++
	}
	If ($checkval3.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath3 -Name $name -Value $value -force;$ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline;write-host "Enabled" -ForegroundColor green -NoNewline;write-host " exists under the RC4 56/128 Registry Key.";$ssl++
	}

# SSL Check totals
	If ($ssl -eq "3") {
		write-host " ";write-host "RC4 " -ForegroundColor yellow -NoNewline;write-host "is completely disabled on this server.";write-host " "
	} 
	If ($ssl -lt "3"){
		write-host " ";write-host "RC4 " -ForegroundColor yellow -NoNewline;write-host "only has $ssl part(s) of 3 disabled.  Please check the registry to manually to add these values";write-host " "
	}
} # End of Disable RC4 function

function DisableSSL3 {
    $TestPath1 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0" -erroraction silentlycontinue
    $TestPath2 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server" -erroraction silentlycontinue
    $registrypath = "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
    $Name = "Enabled"
	$value = "0"
    $checkval1 = Get-Itemproperty -Path "$registrypath" -name $name -erroraction silentlycontinue

# Check for SSL 3.0 Reg Key
	If ($TestPath1 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols", $true)
		$key.CreateSubKey('SSL 3.0')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "SSL 3.0" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

# Check for SSL 3.0\Server Reg Key
	If ($TestPath2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0", $true)
		$key.CreateSubKey('Server')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "SSL 3.0\Servers" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

# Add the enabled value to disable SSL 3.0 Support
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath -Name $name -Value $value -force;$ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline;write-host "Enabled" -ForegroundColor green -NoNewline;write-host " exists under the SSL 3.0\Server Registry Key."
	}
} # End of Disable SSL 3.0 function

Do { 	
	if ($RebootRequired -eq $true){Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL EXCHANGE BEFORE REBOOTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black}
	if ($opt -ne "None") {Write-Host "Last command: "$opt -foregroundcolor Yellow}	
	$opt = Read-Host $menu

	switch ($opt)    {
		1 { #	Prep Mailbox Role - art 1
			Get-ModuleStatus -name ServerManager
        	   	Install-DotNET452
	          	Install-WindowsFeature RSAT-ADDS
			Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
			highperformance
			PowerMgmt
			$RebootRequired = $true
		}
		2 { #	Prep Mailbox Role - Part 2
			Get-ModuleStatus -name ServerManager
            		Install-WinUniComm4
			$RebootRequired = $false
		}
	  	3 {#	Prep Exchange Transport
			Install-windowsfeature ADLDS
			Install-DotNET452
		}
	  	10 {#	Windows Update
			Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
		}
		11 {#	Mailbox Requirement Check
			check-MBXprereq
		}
		12 {#	Edge Transport Requirement Check
			check-EdgePrereq
		}
		20 {#	Install -One-Off - .NET 4.5.2 [MBX or Edge]
			Get-ModuleStatus -name ServerManager
			 Install-DotNET452
		}
		21 {#	Install -One-Off - Windows Features [MBX]
			Get-ModuleStatus -name ServerManager
			Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
		}
		22 {#	Install - One Off - Unified Communications Managed API 4.0
			Install-WinUniComm4
		}
		30 { # Set power plan to High Performance as per Microsoft
			highperformance
		}
		31 { # Disable Power Management for NICs.		
			PowerMgmt
		}
		32 { # Disable SSL 3.0 Support
			DisableSSL3
		}
		33 { # Disable RC4 Support		
			DisableRC4
		}
		98 {#	Exit and restart
			Stop-Transcript
			restart-computer -computername localhost -force
		}
		99 {#	Exit
			if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
				Write-Host "BitsTransfer: Removing..." -NoNewLine
				Remove-Module BitsTransfer
				Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
			}
			popd
			Write-Host "Exiting..."
			Stop-Transcript
		}
		default {Write-Host "You haven't selected any of the available options. "}
	}
} while ($opt -ne 99)