<#-----------------------------------------------------------------------------
Module 'O365_Logon'

Mike O'Neill, Microsoft Senior Premier Field Engineer
http://blogs.technet.microsoft.com/mconeill

Blog post of this module: http://blogs.technet.com/b/mconeill/archive/2015/11/26/o365-powershell-logon-module.aspx

Generated on: 11/6/2015
Version 1.0:    Original post
Version 1.1:    Updated all functions with cmdletbindings
                Updated help content

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

Import-Module O365_Logon -Force
-----------------------------------------------------------------------------#>

#region Request for new credentials function
Function Request-Credential {
<#
.SYNOPSIS
    Requests O365 credentials for error in entering information. 
.DESCRIPTION
    This cmdlet prompts a user to enter in a new user name and password to correct any errors or change sign-in information to an O365 tenant. 
.EXAMPLE
    Request-Credential
.INPUTS
    None
.OUTPUTS
    None
.FUNCTIONALITY
   Prompts for new O365 credentials.
#>
[cmdletbinding()]
param()
        &$Global:UserCredential
} 

#endregion Request for new credentials function

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
.SYNOPSIS
   Creates a PSSession to connect to Compliance Center Online. 
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
[cmdletbinding()]
param()

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
.SYNOPSIS
   Disconnects a PSSession from Compliance Center Online. 
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
[cmdletbinding()]
param()
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
.SYNOPSIS
   Creates a PSSession to connect to Exchange Online. 
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
[cmdletbinding()]
param()

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
.SYNOPSIS
   Disconnects the PSSession from Exchange Online. 
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
[cmdletbinding()]
param()
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
.SYNOPSIS
   Creates a PSSession to connect to SharePoint Online. 
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
[cmdletbinding()]
param()

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
.SYNOPSIS
   Disconnects the PSSession from SharePoint Online. 
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
[cmdletbinding()]
param()
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
.SYNOPSIS
   Creates a PSSession to connect to Skype for Online. 
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
[cmdletbinding()]
param()

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
.SYNOPSIS
   Disconnects the PSSession from Skype for Business Online. 
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
[cmdletbinding()]
param()
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
.SYNOPSIS
   Creates PSSessions to connect to O365 services. 
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
[cmdletbinding()]
param()
    Connect-SPO
    Connect-CCO
    Connect-EXO
    Connect-SfbO     
}

Function Disconnect-O365 {
<#
.SYNOPSIS
   Disconnects the PSSessions from O365 Online. 
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
[cmdletbinding()]
param()
    Disconnect-CCO
    Disconnect-EXO
    Disconnect-SPO
    Disconnect-SfbO
}
#endregion All O365 Sessions connect/disconnect functions

#endregion All Connect/disconnect functions
