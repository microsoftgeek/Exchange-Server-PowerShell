<#-----------------------------------------------------------------------------
Script to look for .NET version on Windows Server 2012 or 2012R2.
Exchange supportability matrix: https://technet.microsoft.com/en-us/library/ff728623(v=exchg.160).aspx
Exchange team blog site about .NET version support: http://blogs.technet.com/b/exchange/archive/2016/02/10/on-net-framework-4-6-1-and-exchange-compatibility.aspx

Mike O'Neill, Microsoft Senior Premier Field Engineer
Main blog page: http://blogs.technet.com/b/mconeill
Blog post about this script: http://blogs.technet.com/b/mconeill/archive/2016/02/27/check-for-net-version-script.aspx
MS Script Center link to download this script: https://gallery.technet.microsoft.com/scriptcenter/Check-for-Net-Version-677f73d6

Generated on: 2/24/2016

Version 1.0 posted: 2/27/2016

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

#region check for PS version

If (($PSVersionTable.PSVersion).Major -le 2.0) {
       Write-Host "`nThis script requires a version of PowerShell 3 or higher, which this is not.`n" -ForegroundColor Red
       Write-Host "PS 3.0: https://www.microsoft.com/en-us/download/details.aspx?id=34595" -ForegroundColor Yellow
       Write-Host "PS 4.0: https://www.microsoft.com/en-us/download/details.aspx?id=40855" -ForegroundColor Yellow
       Write-Host "PS 5.0: https://www.microsoft.com/en-us/download/details.aspx?id=48729" -ForegroundColor Yellow
       Write-Host "Please review the System Requirements to decide which version to install onto your computer.`n" -ForegroundColor Cyan
       Exit
}
#endregion End Check for PS version compatibility

#region Check for OS version
    If ((Get-WMIObject win32_OperatingSystem).Version -notmatch '6.2'){
	    If ((Get-WMIObject win32_OperatingSystem).Version -notmatch '6.3'){
	        Write-Host "`nThis script requires a version of Windows Server: 2012 or 2012 R2, which this is not. Exiting...`n" -ForegroundColor Red
	        Exit
	    }
    }
#endregion

#region Look for .Net version currently installed on OS.

$NetValue = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name "Release"

$NetValue | Select Release, @{
  name="Product"
  expression={
      switch($_.Release) {
        378389 { [Version]"4.5" }
        378675 { [Version]"4.5.1" }
        378758 { [Version]"4.5.1" }
        379893 { [Version]"4.5.2" }
        393295 { [Version]"4.6" }
        393297 { [Version]"4.6" }
        394254 { [Version]"4.6.1" }
        394271 { [Version]"4.6.1" } 
      }
    }
}

	If ($NetValue.Release -gt "379893"){
        Write-Host "The version of .NET installed on this server is not currently (as of 3/1/2016) supported with any version of Exchange." -ForegroundColor Red
}
#endregion

#region 2012 R2 patch numbers to check for

if ((Get-WMIObject win32_OperatingSystem).Version -match '6.3'){

$kb3102467 = Get-HotFix -Id kb3102467 -ErrorAction SilentlyContinue #4.6.1 install update https://support.microsoft.com/en-us/kb/3102436
$kb3083184 = Get-HotFix -Id kb3083184 -ErrorAction SilentlyContinue #4.6.0081 updated security patch (August 11, 2015) https://support.microsoft.com/en-us/kb/3086251 
$kb3045562 = Get-HotFix -Id kb3045562 -ErrorAction SilentlyContinue #4.6.0081 update (September 15, 2015) https://support.microsoft.com/en-us/kb/3045560

    If($kb3102467 -match "kb3102467")
        {
            Write-Host "KB3102467 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3102467 is not installed." -ForegroundColor Green
        }
 
    If($kb3083184 -match "kb3083184")
        {
            Write-Host "KB3083184 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3083184 is not installed." -ForegroundColor Green
        }

    If($kb3045562 -match "kb3045562")
        {
            Write-Host "KB3045562 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3045562 is not installed." -ForegroundColor Green
        }

}
#endregion

#region 2012 patch numbers to check for
if ((Get-WMIObject win32_OperatingSystem).Version -match '6.2'){

$kb3102439 = Get-HotFix -Id kb3102467 -ErrorAction SilentlyContinue #4.6.1 install update https://support.microsoft.com/en-us/kb/3102436
$kb3083185 = Get-HotFix -Id kb3083185 -ErrorAction SilentlyContinue #4.6.0081 updated security patch (August 11, 2015) https://support.microsoft.com/en-us/kb/3086251
$kb3045563 = Get-HotFix -Id kb3045563 -ErrorAction SilentlyContinue #4.6.0081 update (September 15, 2015) https://support.microsoft.com/en-us/kb/3045560

    If($kb3102439 -match "kb3102439")
        {
            Write-Host "KB3102439 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3102439 is not installed." -ForegroundColor Green
        }
 
    If($kb3083185 -match "kb3083185")
        {
            Write-Host "KB3083185 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3083185 is not installed." -ForegroundColor Green
        }

    If($kb3045563 -match "kb3045563")
        {
            Write-Host "KB3045563 is installed, you need to uninstall this patch." -foregroundcolor red
        }
    else
        {
            Write-Host "KB3045563 is not installed." -ForegroundColor Green
        }
}
#endregion
