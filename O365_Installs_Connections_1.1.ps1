
<#-----------------------------------------------------------------------------
O365 install and connection script

Mike O'Neill, Microsoft Senior Premier Field Engineer
http://blogs.technet.com/b/mconeill

Blog post about this script: http://blogs.technet.com/b/mconeill/archive/2015/11/26/o365-installs-connections-ps1.aspx

Generated on: 11/6/2015

LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
-----------------------------------------------------------------------------#>

<#  
.SYNOPSIS
   	Downloads, installs, and has connection options to O365 and online tenant information via PowerShell.

.DESCRIPTION  
    Installs the required modules to access O365 tenant information via PowerShell.

.NOTES  
    Current Version     : 1.1
    
    History				: 1.0 - Posted 11/30/2015 - First iteration
                        : 1.1 - Posted 12/2/13
                            - Fixed error when not running steps in order to see if the 'install' directory exists.
                            - Added logic to O365_Logon module install to confirm when it has completed.
                            - Added logic to the Skype for Busines online module to confirm when it has completed.
                            - Updated parameter statement.
                            - Added logic to the WaaD module to confirm when it has completed.
                            - Added WaaD module dependency for Sign In Assistant.
                            - Added logic to the SharePoint online module to confirm when it has completed.
                            - Added version information for SIA and WaaD installs.
                            
    
    Rights Required		: Local admin on workshop for installing applications
                        : Set-ExecutionPolicy to 'Unrestricted' for the .ps1 file to execute the installs
                        : Requires PowerShell (or ISE) to 'Run as Administrator' to install the applications or modules
                        
    O365 Connect info   : https://technet.microsoft.com/en-us/library/dn568015.aspx
    
.LINK  
    https://gallery.technet.microsoft.com/scriptcenter/O365InstallsConnectionsps1-e4821bf1
    http://blogs.technet.com/b/mconeill/archive/2015/11/29/o365-installs-connections-ps1.aspx

.FUNCTIONALITY
   This script displays options that simplify the process of installing the pre-requisites needed for 
   logging onto individual O365 components: Exchange online, Compliance Center online, SharePoint online and Skype for Business online. 
#>

param(
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true, Mandatory=$false)] 
	[string] $TargetFolder = "C:\Install" ,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
)

#region Detect PS version and 64-bit OS

# Check for 64-bit OS
If($env:PROCESSOR_ARCHITECTURE -match '86') {
        Write-Host "`nThis script only installs 64-bit modules. This machine is not a 64-bit Operating System.`n" -ForegroundColor Red

} # End Check for 64-bit OS

# Check for PowerShell version compatibility
If (($PSVersionTable.PSVersion).Major -lt 3.0) {
       Write-Host "`nThis script requires a version of PowerShell 3 or higher, which this is not.`n" -ForegroundColor Red
       Write-Host "PS 3.0: https://www.microsoft.com/en-us/download/details.aspx?id=34595" -ForegroundColor Yellow
       Write-Host "PS 4.0: https://www.microsoft.com/en-us/download/details.aspx?id=40855" -ForegroundColor Yellow
       Write-Host "PS 5.0: https://www.microsoft.com/en-us/download/details.aspx?id=48729" -ForegroundColor Yellow
       Write-Host "Please review the System Requirements to decide which version to install onto your computer.`n" -ForegroundColor Cyan
       Exit
} # End Check for PowerShell version compatibility

#endregion End Detect PS version and 64-bit OS

Clear-Host

#region Menu display using here string
[string] $menu = @'

	******************************************************************
	                 Logon/Install O365 services
	******************************************************************
	
	Please select an option from the list below:


     1) Log onto all O365 Services
     2) Log onto only Exchange online
     3) Log onto only SharePoint online
     4) Log onto only Skype for Business online
     5) Log onto only Compliance Center online

     10) Launch Windows Update
    	
     11) Install - .NET 4.5.2
     12) Install - MS Online Service Sign-In Assistance for IT Professionals RTW
     13) Install - Windows Azure Active Directory Module
     14) Install - SharePoint Online Module (Reboot required)
     15) Install - Skype for Business Online Module (Reboot required)
     
     20) Install - O365_Logon Module 1.0

     30) Enable PS Remoting on this local computer (Fix WinRM issue)

     31) Launch PowerShell 3.0 download website
     32) Launch PowerShell 4.0 download website
     33) Launch PowerShell 5.0 download website

     90) Launch Blog Post about this script
     91) Launch Blog Post about O365_Logon module
	
     98) Restart this workstation
     99) Exit this script

Select an option.. [1-99]?
'@

#endregion Menu display

#region Installs

Function TestTargetPath { # Test for target path for install temporary directory.
            If ((Test-Path $targetfolder) -eq $true) {
			    Write-Host "Folder: $targetfolder exists." -ForegroundColor Green
		    } 
            Else{
			    Write-Host "Folder: $targetfolder does not exist, creating..." -NoNewline
			    New-Item $targetfolder -type Directory | Out-Null
			    Write-Host "created!" -ForegroundColor Green
            } 
}# End Test for target path for install temporary directory.

function Install-DotNET452{
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if ($val.Release -lt "379893") {
    		GetIt "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
	    	Set-Location $targetfolder
    		[string]$expression = ".\NDP452-KB2901907-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $targetfolder\DotNET452.log"
	    	Write-Host "File: NDP452-KB2901907-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
    		Invoke-Expression $expression
    		Start-Sleep -Seconds 20
    		Write-Host "`n.NET 4.5.2 should be installed by now." -Foregroundcolor Yellow
	} else {
    		Write-Host "`n.NET 4.5.2 already installed." -Foregroundcolor Green
    }
} # end Install .NET 4.5.2

#region Install Windows Azure Active Directory module

Function Check-SIA_Installed { # Check for Sign In Assistant before WaaD can install
        $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
        If ($CheckForSignInAssistant -eq $true) {
                $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                Write-Host "`Sign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is installed" -Foregroundcolor Green
                Install-WAADModule
        }
        Else {
                Write-Host "Windows Azure Active Directory Module stopping installation...`n" -Foregroundcolor Green 
                Write-Host "`nThe Sign In Assistant needs to be installed before the Windows Azure Active Directory module.`n" -Foregroundcolor Red   
        } 
} # End Check for Sign In Assistant before WaaD can install

Function Install-WAADModule {
            Check-Bits #Confirms if BitsTransfer is running on the local host
            $WAADUrl = "https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-FR.msi"
            Start-BitsTransfer -Source $WAADUrl -Description "Windows Azure Active Directory" -Destination $env:temp -DisplayName "Windows Azure Active Directory"
            Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $WAADUrl -Leaf) /quiet /passive"
            Start-Sleep -Seconds 5
            $LoopError = 1 # Variable to error out the loop
            Do {$CheckForWAAD = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\MSOnline"
                Write-Host "Windows Azure Active Directory Module being installed..." -Foregroundcolor Green
                Start-Sleep -Seconds 10
                $LoopError = $LoopError + 1
            }
            Until ($CheckForWAAD -eq $true -or $LoopError -eq 10)
                Start-Sleep -Seconds 5
                If ($CheckForWAAD -eq $true){
                        $WaaDModuleVersion = (get-item C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\Microsoft.Online.Administration.Automation.PSModule.dll).VersionInfo.FileVersion
                        Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion is now installed." -Foregroundcolor Green  
                }
                Else {
                        Write-Host "`nAn error may have occured. Windows Azure Active Directory online module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
                }
}

Function Install-WindowsAADModule {
        $CheckForWAAD = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\MSOnline"
        If ($CheckForWAAD -eq $false){
            Write-Host "`nWindows Azure Active Directory Module starting installation...`n" -Foregroundcolor Green
            Check-SIA_Installed
        }
        Else {
            $WaaDModuleVersion = (get-item C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\Microsoft.Online.Administration.Automation.PSModule.dll).VersionInfo.FileVersion
            If ($WaaDModuleVersion -ge "1.0.8070.2"){
                Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion already installed." -Foregroundcolor Green
            }
            Else {
                Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion already installed." -Foregroundcolor Green
                Write-Host "However, there is a newer version available for download." -ForegroundColor Yellow
                Write-Host "You will need to uninstall your current version and re-install a newer version." -ForegroundColor Yellow
            }            
        }
} #endregion Install Windows Azure Active Directory module

#region Install Sign in Assistant (SIA)
Function Install-SIA {
              Check-Bits #Confirms if BitsTransfer is running on the local host
              $MsolUrl = "http://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi"
              Start-BitsTransfer -Source $MsolUrl -Description "Microsoft Online services" -Destination $env:temp -DisplayName "Microsoft Online Services"
              Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $MsolUrl -Leaf) /quiet /passive"
              Start-Sleep -Seconds 10
              $LoopError = 1 # Variable to error out the loop
              Do {$CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                    Write-Host "Sign In Assistant being installed..." -Foregroundcolor Green
                    Start-Sleep -Seconds 10
                    $LoopError = $LoopError + 1
              }
              Until ($CheckForSignInAssistant -eq $true -or $LoopError -eq 10)
                    Start-Sleep -Seconds 10
                    If ($CheckForSignInAssistant -eq $true){
                            $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                            Write-Host "`nSign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is now installed." -Foregroundcolor Green  
                                       
                    }
                Else {
                        Write-Host "`nAn error may have occured. The Sign In Assistant could be installed or still installing. Rerun this step to confirm." -ForegroundColor Red
                }
                    
                    
}

Function Install-SignInAssistant {
    $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
        If ($CheckForSignInAssistant -eq $false) {
        Write-Host "`nSign In Assistant starting installation...`n" -Foregroundcolor Green
        Install-SIA
        }
            Else {
            $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                If ($SignInAssistantVersion.MSOIDCRLVersion -lt "7.250.4551.0") {
                    Write-Host "`nSign In Assistant starting installation...`n" -Foregroundcolor Green
                    Install-SIA
                }
                Else {
                    Write-Host "`nSign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is already installed" -Foregroundcolor Green
                }
        }
} #endregion Install Sign in Assistant

#region Install Skype for Business Module
Function Install-SfbOModule {
            GetIt "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe"
            Set-Location $targetfolder
            [string]$expression = ".\SkypeOnlinePowershell.exe /quiet /norestart /l* $targetfolder\SkypeOnlinePowerShell.log"
            Write-Host "Skype for Business online starting installation...`n" -NoNewLine -ForegroundColor Green
            Invoke-Expression $expression
    		Start-Sleep -Seconds 5
            $LoopError = 1 # Variable to error out the loop
            Do {$CheckForSfbO = Test-Path "$env:ProgramFiles\Common Files\Skype for business Online\Modules"
                Write-Host "Skype for Business online module being installed..." -Foregroundcolor Green
                Start-Sleep -Seconds 15
                $LoopError = $LoopError + 1
            }
            Until ($CheckForSfbO -eq $true -or $LoopError -eq 10)
    		    If ($CheckForSfbO -eq $true){
                    Start-Sleep -Seconds 10
                    If ($CheckForSfbO -eq $True) {
                        Write-Host "Skype for Business online module now installed." -Foregroundcolor Green
                    }
                    Else {                {
                        Write-Host "`nAn error may have occured. Skype for Business online module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
                }
                Write-Host "             Reboot eventually needed before this module will work.               " -BackgroundColor Red -ForegroundColor Black
            }
            }
}

Function Install-SfbO {
        $CheckForSfbO = Test-Path "$env:ProgramFiles\Common Files\Skype for business Online\Modules"
        If ($CheckForSfbO -eq $false){
            Install-SfboModule
        }
        Else {
            Write-Host "`nSkype for Business Online Module already installed`n" -Foregroundcolor Green
        }
 } #endregion Install Skype for Business Module

#region Install SharePoint Module
Function Install-SPOModule {
              Check-Bits #Confirms if BitsTransfer is running on the local host
              $MsolUrl = "http://blogs.technet.com/cfs-filesystemfile.ashx/__key/telligent-evolution-components-attachments/01-9846-00-00-03-65-75-65/sharepointonlinemanagementshell_5F00_4613_2D00_1211_5F00_x64_5F00_en_2D00_us.msi"
              Start-BitsTransfer -Source $MsolUrl -Description "SharePoint Online Module" -Destination $env:temp -DisplayName "SharePoint Online Module"
              Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $MsolUrl -Leaf) /quiet /passive"
              
              #Logic to confirm that the file downloaded to local client. If not, then launch to website for manual download.
                    $CheckForSPOFileDownload = Test-Path "$env:temp\sharepointonlinemanagementshell_5F00_4613_2D00_1211_5F00_x64_5F00_en_2D00_us.msi"
                    If ($CheckForSPOFileDownload -eq $false) { # Install calls download website for install if download file does not exist
                        Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=35588"
                    }
                    Else {
                        Start-Sleep -Seconds 5
                        $LoopError = 1 # Variable to error out the loop
                        Do {$CheckForSPO = Test-Path "$env:ProgramFiles\SharePoint Online Management Shell"
                            Write-Host "SharePoint Online module being installed..." -Foregroundcolor Green
                            Start-Sleep -Seconds 5
                            $LoopError = $LoopError + 1
                        }
                        Until ($CheckForSPO -eq $true -or $LoopError -eq 10)
                            Start-Sleep -Seconds 10
                            If ($CheckForSPO -eq $true){
                            Write-Host "`nSharePoint online module installation is now complete." -Foregroundcolor Green 
                            }
                            Else {
                                Write-Host "`nAn error may have occured. SharePoint online module could be installed. Rerun this step to confirm." -ForegroundColor Red
                            }
                    }
}

Function Install-SPO {
        $CheckForSPO = Test-Path "$env:ProgramFiles\SharePoint Online Management Shell"
        If ($CheckForSPO -eq $false){
             Install-SPOModule
             Write-Host "             Reboot eventually needed before this module will work.             " -BackgroundColor Red -ForegroundColor Black
        }
        Else {
             Write-Host "`nSharePoint Online Module already installed." -Foregroundcolor Green
        }

} #endregion Install SharePoint Module

#region Install O365_Logon Module

# O365_Logon Module download and extraction
Function Install-O365_LogonModule {  
        $url = "https://gallery.technet.microsoft.com/scriptcenter/O365Logon-Module-a1d9baf2/file/145181/4/O365_Logon.zip" # MS Script Center location
        $output = $env:TEMP
        Import-Module BitsTransfer  
        Start-BitsTransfer -Source $url -Destination $output

function Expand-ZIPFile($file, $destination) {
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
} 

Expand-ZIPFile –File “$output\O365_Logon.zip” –Destination “$env:windir\System32\WindowsPowerShell\v1.0\Modules\”
        Start-Sleep -Seconds 2
        $LoopError = 1 # Variable to error out the loop
        Do {$CheckForO365_LogonModule = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\O365_Logon"
            Write-Host "O365_Logon Module being installed..." -Foregroundcolor Green
            Start-Sleep -Seconds 2
            $LoopError = $LoopError + 1
        } 
        Until ($CheckForO365_LogonModule -eq $true -or $LoopError -eq 10)
            Start-Sleep -Seconds 2
            If ($CheckForO365_LogonModule -eq $true){
                Write-Host "`nO365_Logon Module now installed." -Foregroundcolor Green 
            }
            Else {
                Write-Host "`nAn error may have occured. O365_Logon module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
            }
            
} # End O365_Logon Module download and extraction

# Install O365_Logon module logic to check if already installed or needs to install 
Function Install-O365_Logon {
                If (((Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\O365_Logon") -or (Test-Path "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\O365_Logon") -or (Test-Path "$env:ProgramFiles\WindowsPowerShell\Modules\O365_Logon")) -eq $true){
                        Write-Host "`nO365_Logon Module already installed." -Foregroundcolor Green    
                }      
                Else {
                        Write-Host "O365_Logon Module starting installaion...`n" -Foregroundcolor Green
                        Install-O365_LogonModule  
                }  # End Install O365_Logon module logic to check if already installed or needs to install
} #endregion End Install O365_Logon Module

# Get-Bits function
Function Check-Bits{
    if ((Get-Module BitsTransfer) -eq $null){
			Write-Host "BitsTransfer: Installing..." -NoNewLine
			Import-Module BitsTransfer	
			Write-Host "Installed." -ForegroundColor Green
		}
} # End Get-Bits Function

# GetIt Module
function GetIt ([string]$sourcefile)	{
	if ($HasInternetAccess){
		Check-Bits # check if BitsTransfer is installed
		[string] $targetfile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1) 
		TestTargetPath # Function to confirn or create the $Targetpath download for installable file
		if (Test-Path "$targetfolder\$targetfile"){
			Write-Host "File: $targetfile exists."
		}else{	
			Write-Host "File: $targetfile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$targetfolder\$targetfile"
			Write-Host "Downloaded." -ForegroundColor Green
		}
	}else{
		Write-Host "Internet Access not detected. Please resolve and try again." -foregroundcolor red
	}
} # End GetIt Module function

# Unzip function
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
} # End UnZipIt Function

# New-FileDownload function
function New-FileDownload {
	param (
	[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No source file specified")] 
	[string]$SourceFile,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination folder specified")] 
    [string]$DestFolder,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination file specified")] 
    [string]$DestFile
	)
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
} # End New-FileDownload funtion

# Function Get-ModuleStatus
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
} # End Function Get-ModuleStatus function

#endregion Installs

#region Common used scriptblocks

# Enter Credentials to log onto tenant scriptblock
        $Global:UserCredential = {
            $Global:Credential = Get-Credential -Message "Your logon is your e-mail address for O365."
        } # End Enter Credentials to log onto tenant scriptblock

# Start logic to confirm if logged on user has access to MS online module. Then present users' information as confirmation. 
        $Global:MSolUserScriptBlock = {
            $Global:MSolUser =  Get-MsolUser -UserPrincipalName ($Credential.UserName) #User variable used if logging on user has not mailbox. This confrims that MS Online module is connected
                If ($MSolUser -eq $null){
                        Write-Host "`nYou are not logged into Azure Active Dirctory.`n" -ForegroundColor Red
                }
                Else {
                        Write-Host "`nHello $($MSolUser.DisplayName), you are now logged onto Azure Active Directory.`n" -ForegroundColor Green
                }
 } # End User variable used if logging on user has access to MSOnline. This confrims that MS Online module is connected
 
#endregion Common used scriptblocks

#region All Connect/disconnect functions

#region Compliance Center Online connect/disconnect session functions
Function Connect-CCO {
<#
.Synopsis
   Creates a PSSession to connect to Compliance Center Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet combines all of the steps to successfully log onto an O365 Compliance Center online tenant.
    
.EXAMPLE
   Connect-CCO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Credentials are requested for an O365 tenant. A PSSession is then created using URL's to the O365 tenant, credentials are passed, remote module imported, and a confirmation that a connection is successful is displayed to the end user. 
#>

If ($Credential -eq $null)
{
    &$Global:UserCredential
}
    #Compliance Center Online connect session commands.
    $Global:ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $ccSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking

        # Displaying from online read, last log available time.
        $TenantAdmins = (Get-RoleGroup | Where-Object {$_.name -eq "tenantadmins"})
       If ($TenantAdmins -eq $null){       
            Write-Host "`nYou are not currently connected to the online Compliance Center.`n" -ForegroundColor Red
       }
       Else {
            Write-Host  "`nYou are currently logged into the online Compliance Center.`n" -ForegroundColor Green
            # End Displaying from online read, current tenant version
       }
} 

Function Disconnect-CCO {
<#
.Synopsis
   Disconnects a PSSession from Compliance Center Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet disconnects the log session to an O365 Compliance Center online tenant.
    
.EXAMPLE
   Disconnect-CCO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Terminates the PSSession that is connected to an O365 Compliance Center online remote session and clears the credential variable. 
#>
    # Logic to confirm if the Compliance Center online session has been disconnected.
    If ($ccSession -eq $null) {
        Write-Host "`nThe Compliance Center online session does not exist." -ForegroundColor Yellow
    }
    Else {
        Remove-PSSession $ccSession -EA SilentlyContinue
        $Global:Credential = $null
        If ($ccSession.State -eq "Closed") {
            Write-Host "`nThe Compliance Center online session is now closed." -ForegroundColor Cyan
        }
        Else {
            Write-Host "`nThe Compliance Center online session has not closed." -ForegroundColor Yellow
        }
    }
}
#endregion Compliance Center Online connect/disconnect session functions

#region Exchange Online connect/disconnect session functions

Function Connect-EXO {
<#
.Synopsis
   Creates a PSSession to connect to Exchange Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet combines all of the steps to successfully log onto an O365 Exchange online tenant.
    
.EXAMPLE
   Connect-EXO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Credentials are requested for an O365 tenant. A PSSession is then created using URL's to the O365 tenant, credentials are passed, remote module imported, and a confirmation that a connection is successful is displayed to the end user. 
#>

If ($Credential -eq $null)
{
    &$Global:UserCredential
}
    #Exchange Online connect session commands.
    $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/PowerShell/" -Credential $Credential -Authentication basic -AllowRedirection
    Import-Module (Import-PSSession $EXOSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking
    Connect-MsolService -Credential $Credential
        
        # Logic to confirm tenant access is available.
        &$Global:MSolUserScriptBlock # Check access to MSonline information. 

        # Displaying from online read, current tenant version.
        $AdminDisplayVersion = (Get-OrganizationConfig).AdminDisplayVersion
       If ($AdminDisplayVersion -eq $null){       
            Write-Host "`nYou are not currently connected to an Exchange Online Tenant.`n" -ForegroundColor Red
       }
       Else {
            $EXOVersion = $AdminDisplayVersion -replace "0.20",""
            Write-Host "`nYou are currently logged into Exchange online and your tenant version is:$EXOVersion" -ForegroundColor Green
            # End Displaying from online read, current tenant version
              
            #Confirming connection to Mailbox or displaying no mailbox currently assigned to logged on user
            $myMBX = Get-Mailbox ($Credential.UserName) -ErrorAction "SilentlyContinue" #User Variable to confirm if Exchange Online is available. Checks if mailbox exists for logged in user, then presents name and affirmation that the mailbox is available.
        }
       If ($myMBX -ne $null) { 
            Write-Host -ForegroundColor "Green" "`nHello $($myMBX), you are logged into Exchange Online.`n"
        }
       Else {
            Write-Host -ForegroundColor "Yellow" "`nYour account does not currently have an Exchange Online Mailbox.`n"
    } # End logic check to confirm logged on user

}

Function Disconnect-EXO {
<#
.Synopsis
   Disconnects the PSSession from Exchange Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet disconnects the logged in session to an O365 Exchange online tenant.
    
.EXAMPLE
   Disconnect-EXO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Terminates the PSSession that is connected to an O365 Exchange online remote session and clears the credential variable. 
#>
    # Logic to confirm if the Exchange online session has been disconnected.
    If ($EXOsession -eq $null) {
        Write-Host "`nThe Exchange online session does not exist." -ForegroundColor Yellow
    }
    Else {
        Remove-PSSession $EXOsession
        $Global:Credential = $null
        If ($EXOsession.State -eq "Closed") {
            Write-Host "`nThe Exchange online session is now closed." -ForegroundColor Cyan
        }
        ElseIf ($EXOsession.state -eq "Open") {
            Write-Host "`nThe Exchange online session has not closed." -ForegroundColor Yellow
        }
    }
}
#endregion Exchange Online connect/disconnect session functions

#region SharePoint Online connect/disconnect session functions

 Function Connect-SPO {
<#
.Synopsis
   Creates a PSSession to connect to SharePoint Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet combines all of the steps to successfully log onto an O365 SharePoint online tenant.
    
.EXAMPLE
   Connect-SPO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Credentials are requested for an O365 tenant. A PSSession is then created using URL's to the O365 tenant, credentials are passed, remote module imported, and a confirmation that a connection is successful is displayed to the end user. 
#>

If ($Credential -eq $null)
{
    &$Global:UserCredential
}
    Write-Host 'You only need the "Tenant" name. http://<Tenant>-admin.sharepoint.com' -ForegroundColor Magenta
    Write-Host 'This script takes the input tenant name and connects to the SharePoint admin site.' -ForegroundColor Magenta
    Write-Host 'What is the SharePoint Domain Host name you want to connect to? ' -ForegroundColor Yellow -NoNewline    
    $DomainHost = Read-Host 
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    Connect-SPOService -Url https://$DomainHost-admin.sharepoint.com -credential $credential

      # Displaying from online read, last log available time.
        $SPOTenantLogLastAvailable = Get-SPOTenantLogLastAvailableTimeInUtc
       If ($SPOTenantLogLastAvailable -eq $null){       
            Write-Host "`nYou are currently not connected to a SharePoint Online Tenant.`n" -ForegroundColor Red
       }
       Else {
            Write-Host "`nYou are currently logged into SharePoint online and the tenant log last availabile time (UTC) is: $SPOTenantLogLastAvailable`n" -ForegroundColor Green
            # End Displaying from online read, current tenant version
       }
}

Function Disconnect-SPO {
<#
.Synopsis
   Disconnects the PSSession from SharePoint Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet disconnects the logged in session to an O365 SharePoint online tenant.
    
.EXAMPLE
   Disconnect-SPO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Terminates the PSSession that is connected to an O365 Exchange online remote session and clears the credential variable. 
#>
    # Logic to confirm if SharePoint online session has been disconnected
    $ErrorActionPreference = "SilentlyContinue"
    $SPOTenant = Get-SPOTenant
    If ($SPOTenant -eq $null){       
            Write-Host "`nThe SharePoint online session does not exist." -ForegroundColor Yellow
       }
       Else {
            $ErrorActionPreference = "SilentlyContinue"
            Disconnect-SPOService
            $Global:Credential = $null
            Write-Host "`nThe SharePoint online connection has been closed." -ForegroundColor Cyan
       }
}
#endregion SharePoint Online connect/disconnect session functions

#region Skype for Business Online connect/disconnect session functions98
Function Connect-SfbO {
<#
.Synopsis
   Creates a PSSession to connect to Skype for Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet combines all of the steps to successfully log onto an O365 Skype for Business online tenant.
    
.EXAMPLE
   Connect-SfbO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Credentials are requested for an O365 tenant. A PSSession is then created using URL's to the O365 tenant, credentials are passed, remote module imported, and a confirmation that a connection is successful is displayed to the end user. 
#>

If ($Credential -eq $null)
{
    &$Global:UserCredential
}
    Import-Module LyncOnlineConnector
    $global:sfboSession = New-CsOnlineSession -Credential $credential
    Import-Module (Import-PSSession $sfboSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking

    # Logic to confirm tenant access is available.
          # Displaying from online read, last log available time.
        $CSTenant = (Get-CsTenant).DisplayName
       If ($CSTenant -eq $null){       
            Write-Host "`nYou are currently not connected to a Skype for Business Online Tenant.`n" -ForegroundColor Red
       }
       Else {
            Write-Host "`nYou are currently logged into Skype for Business online and the tenant display name is: $CSTenant`n" -ForegroundColor Green
            # End Displaying from online read, current tenant version
            }  # End Logic to confirm tenant access is available.
}

Function Disconnect-SfbO {
<#
.Synopsis
   Disconnects the PSSession from Skype for Business Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet disconnects the logged in session to an O365 Skype for Business online tenant.
    
.EXAMPLE
   Disconnect-SfbO

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Terminates the PSSession that is connected to an O365 Exchange online remote session and clears the credential variable. 
#>
    # Logic to confirm if the Skype for Business online session has been disconnected.
    If ($sfboSession -eq $null) {
        Write-Host "`nThe Skype for Busienss online session does not exist." -ForegroundColor Yellow
    }
    Else {   
        Remove-PSSession $sfboSession -EA SilentlyContinue
        $Global:Credential = $null
        If ($sfboSession.State -eq "Closed") {
            Write-Host "`nThe Skype for Business online session is now closed." -ForegroundColor Cyan
        }
           Else {
            Write-Host "`nThe Skype for Business online session has not closed." -ForegroundColor Yellow
        }
    }
}
#endregion End Skype for Business Online connect/disconnect session functions

#region All O365 Sessions connect/disconnect session functions

Function Connect-O365 {
<#
.Synopsis
   Creates PSSessions to connect to O365 services. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet combines all of the steps to successfully log onto an O365 tenant.
    
.EXAMPLE
   Connect-O365

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Credentials are requested for an O365 tenant. A PSSession is then created using URL's to the O365 tenant, credentials are passed, remote module imported, and a confirmation that a connection is successful is displayed to the end user. 
#>
    Connect-SPO
    Connect-CCO
    Connect-EXO
    Connect-SfbO     
}

Function Disconnect-O365 {
<#
.Synopsis
   Disconnects the PSSessions from O365 Online. 

.SYNTAX
    None

.DESCRIPTION
   This cmdlet disconnects the logged in sessions to an O365 tenant.
    
.EXAMPLE
   Disconnect-O365

.INPUTS
    None

.OUTPUTS
    None
    
.FUNCTIONALITY
   Terminates the PSSession that is connected to an O365 Exchange online remote session and clears the credential variable. 
#>
    Disconnect-CCO
    Disconnect-EXO
    Disconnect-SPO
    Disconnect-SfbO
}
#endregion All O365 Sessions connect/disconnect functions

#endregion All Connect/disconnect functions

#region Menu action

Do { 	
	if ($opt -ne "None") {Write-Host "Last command: "$opt -foregroundcolor Yellow}	
	$opt = Read-Host $menu

	switch ($opt)    {
    			
	  	1 { # Log onto all services
            Connect-SPO
            Connect-CCO
            Connect-EXO
            Connect-SfbO
            Exit
        }

        2 { # Log onto Exchange Online
            Connect-EXO
            Exit
        }

        3 { # Log onto SharePoint Online
            Connect-SPO
            Exit
        }

        4 { # Log onto Skype for Business Online
            Connect-SfbO
            Exit
        }

        5 { #Log onto Compliance Center Online
            Connect-CCO
            Exit
        }

	  	10 { # Windows Update
			Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
		}
		
		11 { # Install - .NET 4.5.2
			 Install-DotNET452
		}

        12 { # Install MS Online Service Sign-In Assistance for IT Professionals RTW
            Install-SignInAssistant
        }

        13 { # Install - Windows Azure Active Directory Module for Windows PowerShell (64-bit)
            Install-WindowsAADModule
        }

        14 { # Install - SharePoint Online Module
            Install-SPO
        }

        15 { # Install - Skype for Business Online Module
            Install-SfbO
        }
         
        20 { # Install - O365_Logon Module 1.0
            Install-O365_Logon
        }

        30 { # Enable PS Remoting. This fixes WinRM error
            Enable-PSRemoting -Force -SkipNetworkProfileCheck
        }

        31 { # Launches PS 3.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=34595"
        }

        32 { # Launches PS 4.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=40855"
        }

        33 { # Launches PS 5.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=48729"
        }

        90 { # Launches Blog Post about this script that includes instructions
            Start-Process "http://blogs.technet.com/b/mconeill/archive/2015/11/26/o365-installs-connections-ps1.aspx"
        }
       
        91 { # Launches Blog Post about the O365_Logon module that includes instructions
            Start-Process "http://blogs.technet.com/b/mconeill/archive/2015/11/27/o365-powershell-logon-module.aspx"
        }

        92 {"This is cool"}

		98 { # Exit and restart
			Restart-Computer -computername localhost -force
		}

		99 { # Exit
			if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
				Write-Host "BitsTransfer: Removing..." -NoNewLine
				Remove-Module BitsTransfer
				Write-Host "Removed." -ForegroundColor Green
			}
			Write-Host "Exiting..."
		}
		
        default {Write-Host "You haven't selected any of the available options."}
	}
} while ($opt -ne 99)

#endregion Menu action