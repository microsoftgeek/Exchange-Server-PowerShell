<#-----------------------------------------------------------------------------
Exchange online aggressive termination script

Mike O'Neill, Microsoft Senior Premier Field Engineer
http://blogs.technet.microsoft.com/mconeill

Blog post about how to use this script: https://blogs.technet.microsoft.com/mconeill/2016/08/05/exchange-online-…rmination-script/

Generated on: 6/1/2016

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
   	Steps to take when an employee is aggressively terminated.

.DESCRIPTION  
    Exchange online caching process can take minutes or several hours to replicate throught the entire O365 environment.
    When a user is aggressively terminated, there are several actions that you should perform to minimize access to the tenant.
    This script disables several key elements of a users' access to the O365 tenant. Thiese steps can also be used for any 
    user deletion process. 

.NOTES  
    Current Version     : 1.0 release
    
    History				: 8/4/2016 Initial publish to internet
    
    Rights Required		: Exchange tenant recipient management role
                       
    
.LINK  
    https://gallery.technet.microsoft.com/Exchange-Online-Aggressive-fc144a91

.FUNCTIONALITY
   This script helps companies block and deny access to Exchange online for terminated employees. 
#>

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

#region Select end user to be aggressively terminated
Function Input-MailboxAlias {
    Clear-Host
    Write-Host "Input the user's alias you wish to terminate: " -ForegroundColor Yellow -NoNewline    
    $Global:DisplayName = Read-Host
}
#endregion

#region Present User's information about to be modified
Function Write-UserInformation {
    Write-Host "The user alias $Global:DisplayName settings are currently:" -ForegroundColor Green -NoNewline
    
    #Display current license values of selected user
    $Global:OnlineUser = (Get-Mailbox $Global:DisplayName).UserPrincipalName # Converts mailbox alias to UPN of MsolUser
    $Global:LicenseAccountSkuId = (Get-MsolAccountSku).AccountSkuID # Retrieves tenant account license SKU ID
    
    #(Get-MsolUser -UserPrincipalName $Global:OnlineUser).Licenses.ServiceStatus
    
    #User's MSOL license status
    Get-MsolUser -UserPrincipalName $Global:OnlineUser | select licenses,BlockCredential
    
    #User's mailbox limits
    Get-Mailbox $Global:DisplayName | ft IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,LitigationHoldEnabled
    
    #User's mailbox features
    Get-CASMailbox $Global:DisplayName | ft MAPIEnabled,OWAEnabled,OWAforDevicesEnabled,ActiveSyncEnabled,PopEnabled,ImapEnabled,EWSEnabled

    #User's EAS settings
    Get-CASMailbox $Global:DisplayName | ft ActiveSyncAllowedDeviceIDs,ActiveSyncBlockedDeviceIDs,ActiveSyncEnabled
}
#endregion

#region Test for O365 mailbox
Function Confirm-O365Mailbox {  #Confirming if Mailbox is in O365 or on premises and if there is an O365 license assigned to the end user

$UserMBX = Get-Mailbox $Global:DisplayName -ErrorAction SilentlyContinue #User Variable to confirm if Exchange Online is available. Checks if mailbox exists for logged in user, then presents name and affirmation that the mailbox is available.
        
    If ($UserMBX -ne $null) { 

            #Need to check for 'onmicrosoft.com address'
            $smtpAddress = ($UserMBX).EmailAddresses
            
            If ($smtpAddress -match "onmicrosoft"){
                Write-Host -ForegroundColor Green "`nThe user alias: $($Global:DisplayName), has an Exchange Online mailbox.`n"
            }
            Else{
                    If ($Global:LicenseAccountSkuId -eq $null) #Need to check for O365 license
                    {
                        Write-Host "The user alias: $($Global:DisplayName), does not currently have an O365 license.`n" -ForegroundColor Red
                        Break
                    }
                    Else{
                            Write-Host "The user alias: $($Global:DisplayName), has an O365 license but there is no '.onmicrosoft.com' SMTP address listed for the user." -ForegroundColor Yellow
                        }
                }

            }
       
       Else {
            Write-Host -ForegroundColor Red "`nThere is no licensed mailbox found for the user alias: $($Global:DisplayName).`n"
            Break

    } # End logic check to confirm logged on user 
}
#endregion

#region confirm to act upon the user that was sent to the input
Function Confirm-UpdateUserSettings {
$title = "Disable $Global:DisplayName O365 settings."
$message = "Are you sure you want to disable $Global:DisplayName's account?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Disabling $Global:DisplayName's account."


$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Returning to PowerShell prompt."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

switch ($result)
    {
        0 {"You selected Yes. Processing..."
            Enable-LitigationHold
            Set-MailboxLimits
            Set-CASMailboxServices
            Remove-O365Access
            Remove-EASDevices
            Write-UserInformation
          }
        1 {"You selected No. Exiting Script."
            Break
          }
    }
}
#endregion

#region Place users' mailbox on litigation hold
Function Enable-LitigationHold {
    Set-Mailbox $Global:DisplayName -LitigationHoldEnabled $true -WarningAction SilentlyContinue #Sets the users' mailbox of litigation hold indefinetly.
}

<# Undo litigation hold of user for testing
Set-Mailbox $Global:DisplayName -LitigationHoldEnabled $false
#>

#endregion

#region Set users' mailbox to not be able to send messages
Function Set-MailboxLimits {
    Set-Mailbox $Global:DisplayName -IssueWarningQuota 0KB -WarningAction SilentlyContinue #Sets issue warning quota to 0k message size
    Set-Mailbox $Global:DisplayName -ProhibitSendQuota 0KB -WarningAction SilentlyContinue #Sets sending quota to 0k message size
    Set-Mailbox $Global:DisplayName -ProhibitSendReceiveQuota 0KB -WarningAction SilentlyContinue #Sets send/reveive quota to 0k message size
}

<#undo for resetting user during testing
Set-Mailbox $Global:DisplayName -ProhibitSendReceiveQuota 50gb 
Set-Mailbox $Global:DisplayName -ProhibitSendQuota 49.5gb 
Set-Mailbox $Global:DisplayName -IssueWarningQuota 49gb
#>

#endregion

#region Disable Mailbox features
Function Set-CASMailboxServices {
    Set-CASMailbox $Global:DisplayName -MAPIEnabled $False -WarningAction SilentlyContinue # Disables the MAPI connection to the mailbox
    Set-CASMailbox $Global:DisplayName -OWAEnabled $False -WarningAction SilentlyContinue # OWA disabled for mailbox
    Set-CASMailbox $Global:DisplayName -OWAforDevicesEnabled $False -WarningAction SilentlyContinue # OWA disabled for mailbox
    Set-CASMailbox $Global:DisplayName -ActiveSyncEnabled $False -WarningAction SilentlyContinue # Disable EAS
    Set-CASMailbox $Global:DisplayName -PopEnabled $False -WarningAction SilentlyContinue # Disable POP3
    Set-CASMailbox $Global:DisplayName -ImapEnabled $False -WarningAction SilentlyContinue # Disable IMAP4
    Set-CASMailbox $Global:DisplayName -EWSEnabled $False -WarningAction SilentlyContinue # Disable EWS
}

<#undo for resetting user during testing
Set-CASMailbox $Global:DisplayName -MAPIEnabled $True 
Set-CASMailbox $Global:DisplayName -OWAEnabled $True
Set-CASMailbox $Global:DisplayName -OWAforDevicesEnabled $True 
Set-CASMailbox $Global:DisplayName -ActiveSyncEnabled $True 
Set-CASMailbox $Global:DisplayName -PopEnabled $True 
Set-CASMailbox $Global:DisplayName -ImapEnabled $True 
Set-CASMailbox $Global:DisplayName -EWSEnabled $True
#>

#endregion

#region Remove tenant license from user
Function Remove-O365Access {
    Set-MsolUserLicense -UserPrincipalName $Global:OnlineUser -RemoveLicenses $LicenseAccountSkuId # Removes license from AAD user.
    Set-MsolUser -UserPrincipalName $Global:OnlineUser -BlockCredential $true # Blocks access to online resources.
}

<#undo removing license for testing
Set-MsolUserLicense -UserPrincipalName $Global:OnlineUser -AddLicenses $LicenseAccountSkuId
Set-MsolUser -UserPrincipalName $Global:OnlineUser -BlockCredential $false
Get-MsolUser -UserPrincipalName $Global:OnlineUser | fl BlockCredential,*license*
#>

#endregion

#region Block current known EAS devices for user
Function Remove-EASDevices {
    $EASDevices = Get-CASMailbox $Global:DisplayName
    Set-CASMailbox $Global:DisplayName -ActiveSyncBlockedDeviceIDs $EASDevices.ActiveSyncAllowedDeviceIDs
    Set-CASMailbox $Global:DisplayName -ActiveSyncAllowedDeviceIDs $null
    Set-CASMailbox $Global:DisplayName -ActiveSyncEnabled $false
}

<# Undo for resetting user during testing
Set-CASMailbox $Global:DisplayName -ActiveSyncAllowedDeviceIDs $EASDevices.ActiveSyncBlockedDeviceIDs
Set-CASMailbox $Global:DisplayName -ActiveSyncBlockedDeviceIDs $null
Set-CASMailbox $Global:DisplayName -ActiveSyncAllowedDeviceIDs $null
Set-CASMailbox $Global:DisplayName -ActiveSyncEnabled $true
#>

#endregion

#region Steps to run for script.
    Input-MailboxAlias
    Write-UserInformation
    Confirm-O365Mailbox
    Confirm-UpdateUserSettings
#endregion
