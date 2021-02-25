
#region Do not run entire file
#This is to stop F5 from being run in ISE or running the file as a ps1 script.
Write-host "When using this file in ISE, only run lines using F8." -ForegroundColor Yellow
Write-Host "Do not run this entire file in PowerShell and/or hit F5 when using ISE." -ForegroundColor Yellow
Break
#endregion

<# 
.DESCRIPTION 
    Workshop demo files. 
 Mike O'Neill, Microsoft Senior Premier Field Engineer
    Main blog page: http://blogs.technet.microsoft.com/mconeill

LEGAL DISCLAIMER:

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
against any claims or lawsuits, including attorneys' fees, that arise or result
from the use or distribution of the Sample Code.
#> 

#region Check for PS version
function Test-PowerShellVersion {
    [CmdletBinding()]
    [OutputType([Boolean])]
    param
    (
        [Int]
        $MinMajorVersion = 3,

        [Int]
        $MinMinorVersion = 0
    )

    if ($null -eq $PSVersionTable -or $null -eq $PSVersionTable.PSVersion)
    {
        Write-Host "[$([DateTime]::Now)] Failed to detect PowerShell version. This script requires at least PowerShell version $MinMajorVersion.$MinMinorVersion." -ForegroundColor Cyan
        return $true
    }

    if ($PSVersionTable.PSVersion.Major -lt $MinMajorVersion -or ($PSVersionTable.PSVersion.Major -eq $MinMajorVersion -and $PSVersionTable.PSVersion.Minor -lt $MinMinorVersion))
    {
        Write-Host "[$([DateTime]::Now)] This script requires at least PowerShell version $MinMajorVersion.$MinMinorVersion." -ForegroundColor Yellow
        return $false
    }

    Write-Host "[$([DateTime]::Now)] Found acceptable PowerShell version: $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)." -ForegroundColor Green

    return $true
} 
#endregion

#region PowerShell additional information

#PowerShell Team blog site
Start-Process https://blogs.msdn.microsoft.com/powershell/

#PowerShell syle guide
Start-Process https://github.com/PowerShell/DscResources/blob/master/StyleGuidelines.md

#ISE Keyboard shortcuts
Start-Process https://msdn.microsoft.com/en-us/powershell/scripting/core-powershell/ise/keyboard-shortcuts-for-the-windows-powershell-ise 

#Hey scripting guy
Start-Process http://blogs.technet.com/b/heyscriptingguy/

#Windows PS Blog
Start-Process http://blogs.msdn.com/b/powershell/

#PowerShell magazine
Start-proces http://powershell.net/

#PowerShell IDERA community
Start-process http://community.idera.com/powershell

#PowerShell Plus, free download application
start-process https://www.idera.com/productssolutions/freetools/powershellplus

#Free PS eBooks
Start-Process http://blogs.technet.com/b/pstips/archive/2014/05/26/free-powershell-ebooks.aspx

#Free PS cookbooks
Start-Process http://www.powertheshell.com/cookbooks/

#Recommened book by Dan Sheehan:
Start-Process "https://www.amazon.com/PowerShell-Depth-Don-Jones/dp/1617292184/ref=pd_sim_14_5?_encoding=UTF8&pd_rd_i=1617292184&pd_rd_r=R10HR74YWVZ3KGBRDHYE&pd_rd_w=g0kGO&pd_rd_wg=aqx2t&psc=1&refRID=R10HR74YWVZ3KGBRDHYE"

#Future of ISE:
Start-Process https://blogs.msdn.microsoft.com/powershell/2017/05/10/announcing-powershell-for-visual-studio-code-1-0/

#region Random PS stuff

#Get stock information
Function Get-StockPrice {
param ($TickerName = 'msft')

((wget "http://www.nasdaq.com/symbol/$TickerName").AllElements | where id -eq "qwidget_lastsale").innerText
}

#Get Comcast Data
Function Get-ComcastData {
<#
.Synopsis
   Gathers data from Comcast on the current usage of data.
.DESCRIPTION
   Using web access, with password of a valid user of Comcast, this function presents data of home usage for the past month. 
.EXAMPLE
   Get-ComcastData -UserName JonDoe@comcast.net -Password Password1
.INPUTS
   User name and Password required. 
.OUTPUTS
   Provides Comcast monthly data usage.
.NOTES
   Developed by Dan Orum.
#>
[cmdletbinding()]
param ($username="User Name Here",$password = "Password Here")

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
$loginUri="https://login.comcast.net/login?r=comcast.net&s=oauth&continue=https%3A%2F%2Flogin.comcast.net%2Foauth%2Fauthorize%3Fclient_id%3Dmy-account-web%26redirect_uri%3Dhttps%253A%252F%252Fcustomer.xfinity.com%252Foauth%252Fcallback%26response_type%3Dcode%26state%3D%2523%252Fdevices%26response%3D1&client_id=my-account-web"

$R=Invoke-WebRequest -Uri $loginUri -SessionVariable Comcast -Method Get  
$form = $R.Forms["signin"]

$Form.Fields["user"]=$username
$Form.Fields["passwd"]=$password

$R=Invoke-WebRequest -UseBasicParsing -Uri "https://login.comcast.net/login"  -WebSession $Comcast -Method POST -Body $Form.Fields 
$R=Invoke-WebRequest -UseBasicParsing -Uri "https://customer.xfinity.com/apis/services/internet/usage"  -webSession $Comcast -Method Get 

$months=($R.Content | ConvertFrom-Json).usageMonths
$usage=[int]$months[$months.Length-1].homeUsage
$now = Get-Date -format G
Write-Host "$now,$usage GB" 
}

#Order a pizza
Function Get-Pizza {

}

#endregion

#endregion

#region Exchange/O365 demo information

#region Log onto O365 
$Cred = Get-Credential 
$Session = New-PSSession –ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://ps.outlook.com/powershell `
    -Credential $Cred `
    -Authentication Basic `
    -AllowRedirection
Import-PSSession $Session -Prefix EXO
Import-Module msonline
Connect-Msolservice


#region Log into mulitple sessions concurrently
Function Connect-MultipleEXO {

$Cred = Get-Credential
$SessionNumbers = 1,2,3

Foreach ($SessionNumber in $SessionNumbers) {

$Session = New-PSSession –ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://ps.outlook.com/powershell `
    -Credential $Cred `
    -Authentication Basic `
    -AllowRedirection
Import-PSSession $Session -Prefix $SessionNumber -AllowClobber
Import-Module msonline
Connect-Msolservice
}
}
#endregion

#endregion

#region Start-stop transcript
start-transcript c:\temp\PowerShell_transcript.txt
stop-transcript
notepad.exe c:\temp\PowerShell_transcript.txt
#endregion

#region Filter options
Get-Mailbox
 
Get-Mailbox -filter {Department -like "marketing"}

Get-User -filter {Department -like "marketing"}

Get-User -filter {Department -like "marketing"} | Format-List 

Get-User -filter {Department -like "marketing"} | Format-Table name, department, office

Get-User -filter {Department -like "marketing"} | Get-Mailbox

Get-User -filter {Department -like "marketing"} | Set-User -Office "chicago"

Get-User -filter {(Department -like "*marketing*") -AND (RecipientType -eq "UserMailbox")} |ft name, Department, RecipientType

Get-User -filter {(Department -like "*marketing*") -AND (RecipientType -eq "UserMailbox")} | Set-Mailbox -IssueWarningQuota 600mb

Get-Mailbox sally | ft name, IssueWarningQuota

Get-User -filter {(Department -like "*marketing*") -AND (RecipientType -eq "UserMailbox")} |Get-Mailbox | ft name, IssueWarningQuota

Get-User| Where-object {$_.Department -eq "Marketing"}

Get-User -Filter {Department -eq "Marketing"}

#endregion

#region Distribution Groups
New-DistributionGroup -Name "HR and security“  -Alias HR_Security -Type "Distribution"
 
# Get users in a department and add them to a DG:
 
$MRK_USR = Get-User -filter {Department -like "*marketing*"} 
 
$MRK_USR| ForEach-Object {Add-DistributionGroupMember -Identity "HR_Security" -Member $_.name}
 
#Check if it worked:
 
Get-DistributionGroupMember "HR and Security"

Get-DistributionGroupMember HR_Security | Set-User -Office “Chicago”

#Dynamic Distribution Groups
New-DynamicDistributionGroup -Name "Legal Team" -Alias Legal -IncludedRecipients "MailboxUsers,MailContacts"  -ConditionalDepartment “Legal”
Function Get-DDGMembers {
<#
.Synopsis
   Get DDG members.
.DESCRIPTION
   List out members of a given Dynamic Distribution List
.EXAMPLE
   Get-DDGMembers -Group Sales

   Lists all members of the Dynamic Distribution Group called 'Sales'
.EXAMPLE
   Get-DDGMembers -Group 'Legal Team'

   Lists all members of the Dynamic Distribution Group called 'Legal Team'. Note the quotes around a value that includes spaces in the name.
.INPUTS
   None
.OUTPUTS
   Present results from the requested Dynamic Distribution Group
#>
    param($Group)
        $ddg = Get-DynamicDistributionGroup $Group ;Get-Recipient -RecipientPreviewFilter $ddg.RecipientFilter | FT Alias
}
#endregion

#region Send Test messages
#This command is to drop email using SMTP server

$msolcred = Get-Credential #save the credential of from address

Send-MailMessage –From user@domain.com –To user@hotmail.com –Subject “Test Email” –Body “Test SMTP Relay Service” -SmtpServer smtp.office365.com -Credential $msolcred -UseSsl -Port 587
Send-MailMessage –From user@domain.onmicrosoft.com –To user@hotmail.com –Subject “Test Email” –Body “Test SMTP Relay Service” -SmtpServer smtp.office365.com -Credential $msolcred -UseSsl -Port 587

#This command is to send email using MX records

Send-MailMessage –From user@domain.com –To user@hotmail.com –Subject “Test Email” –Body “Test SMTP Relay Service” -SmtpServer domain.mail.protection.outlook.com 

#endregion 

#region Reports

#Tenant information
Start-Process http://lynx.office.net

Get-MailTrafficReport

Get-MailboxActivityReport -ReportType Monthly -StartDate 01/01/2015 -EndDate 02/28/2015 |Out-File c:\temp\mailstats.txt

Get-MailboxUsageDetailReport -StartDate (Get-Date).AddMonths(-1) -EndDate (Get-Date) |Out-File \\someserver\someshare\temp\mailstat(Get-Date).txt

Get-MailboxUsageDetailReport -StartDate 01/01/2015 -EndDate 02/28/2015 |Export-Csv -path c:\temp\mailstats.csv 

Get-MailboxUsageDetailReport -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) |Export-Csv -path c:\temp\mailstats2.csv –notypeinformation

Get-MailDetailDlpPolicyReport -StartDate 01/01/2015 -EndDate 02/28/2015 -SenderAddress  katiej@<tenant>.onmicrosoft.com |Out-File c:\temp\mailstats.txt

#endregion 

#region In-Place Hold - eDiscovery
New-MailboxSearch "Hold-Case" -SourceMailboxes "joe@contoso.com" -InPlaceHoldEnabled $true 

Set-MailboxSearch "Hold-Case" -InPlaceHoldEnabled $false 

Remove-MailboxSearch "Hold-Case"

#In-Place eDiscovery
New-MailboxSearch "Discovery-CaseId012" -StartDate "1/1/2009" -EndDate "12/31/2015" -SourceMailboxes "Official_User" -TargetMailbox "Discovery Search Mailbox" -SearchQuery '"Corrupt" AND "Security"' -MessageTypes Email -IncludeUnsearchableItems -LogLevel Full 


#endregion

#region Audit Logging
Get-Mailbox | ft name,auditenabled
Set-Mailbox -AuditEnabled $true

Get-Mailbox | Set-Mailbox –AuditEnabled $false

Start-Process "https://technet.microsoft.com/en-us/library/ff459237(v=exchg.160).aspx" #List of what is really audited

Get-MailboxPermission bandit | ft user,accessrights
Get-MailboxFolderPermission bandit | ft user,accessrights

#see when MFA ran last
Export-MailboxDiagnosticLogs -extendedproperties -Identity bandit
# MFCMAPI value: PR_LAST_MODIFICATION_TIME, PidTagLastModificationTime, ptagLastModificationTime  

#endregion

#region EXO limits - Safe URL's

#listed limits
Start-Process https://technet.microsoft.com/en-us/library/exchange-online-limits.aspx

#throttle limits
Start-Process https://blogs.msdn.microsoft.com/exchangedev/2011/06/23/exchange-online-throttling-and-limits-faq/
Get-ThrottlingPolicy

#safe URL listing for O365 services
Start-process https://support.office.com/en-us/article/Office-365-URLs-and-IP-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2

#endregion 

#region Networking troubleshooting information

# Peering points connections
Start-Process http://www.peeringdb.com/view.php?asn=8075

# Test for Peering point from current workstation
tracert outlook.office365.com

#Test Connectivity site
Start-Process http://testconnectivity.microsoft.com/

#Hybrid Envrionment Free/busy site
Start-Process http://support.microsoft.com/kb/2555008

#DNS check for autodiscover (NSLookup PS command is listed below)
Resolve-DnsName Autodiscover.mcdeo.onmicrosoft.com

ping www.bing.com
#endregion

#region filtering vs. where-object issue with O365 and throttling process


<#Get-Mailbox cannot handle the -filter. 
The first lines takes hours against O365, is throttled, and eventually is cancelled from O365.
Filter option is preferred in PowerShell but not always available. 
#>

Get-Mailbox | Where-Object {$_.WhenMailboxCreated -ge "4/10/2018"} #works properly but can time out due to throttling
Get-Mailbox | Where-Object {$_.WhenMailboxCreated -ge (get-date).AddDays(-2)} #works properly but can time out due to throttling



Get-Mailbox -filter {whenmailboxcreated -ge "4/26/2018"}  #works with no errors and is the correct information



Get-Mailbox -filter {whenmailboxcreated -gt (get-date).adddays(-2)} #errors out. Filter with date cmdlet not available for this cmdlet



Get-user | Where-Object {$_.whencreated -ge (get-date).(-4)} | Get-Mailbox | Where-Object {$_.whenmailboxcreatd -gt (get-date).(-2) }



$dateNeeded = (Get-Date).AddDays(-2)
$dateNeeded
Get-Mailbox -filter {whenmailboxcreated -gt $dateNeeded}



Get-user | Where-Object {$_.whencreated -le (get-date).(-4)} | Get-Mailbox | Where-Object {$_.whenmailboxcreatd -lt (get-date).(-2)}

#Start-RobustCloudCommand Allows stability and avoids O365 throttling
Start-Process https://gallery.technet.microsoft.com/office/Start-RobustCloudCommand-69fb349e

#endregion 

#region Look for Event ID 64: user deleted meeting request and did not act upon it

Get-EventLog Application

Get-EventLog Application | Where-Object {$_.Source -eq "Outlook"}

Get-EventLog Application | Where-Object {$_.EventID -eq 64}

Get-EventLog Application | Where-Object {($_.Source -eq "Outlook") -and ($_.EventID -eq 64)}

#how to check for tentative meetings. Change 'localhost' to client you are working with. 
Invoke-Command -ComputerName localhost -ScriptBlock {Get-EventLog Application | Where-Object {($_.Source -eq "Outlook") -and ($_.EventID -eq 64)}}

#endregion

#region Messaging
#Get Quarantine Messages
Get-QuarantineMessage -StartReceivedDate (Get-Date).AddDays(-7) -EndReceivedDate 02/14/2013

#To release a quarantined message
Get-QuarantineMessage -MessageID 5c695d7e-6642-4681-a4b0-9e7a86613cb7@contoso.com | Release-QuarantineMessage 

#Search for content to delete from within a mailbox
Get-mailbox | Search-Mailbox  -SearchQuery "Wallmart"  -TargetFolder DeletedFromJohnSmith -TargetMailbox DumpEMailFromMailbox –DeleteContent

#region Search by domain for Exchange on premises
#All messages with recipients like a domain.
Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-date).AddMonths(-1) -End (Get-Date) | Where-Object {$_.recipients -like "*@Contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv \\someserver\someshare\ExchangeLogResults_ToUsersOfContoso$(get-date -f dd-MM-yyyy)services.csv

#All messages with recipients not like a domain.
Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-date).AddMonths(-1) -End (Get-Date) | Where-Object {$_.recipients -notlike "*@Contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\ExchangeLogResults_ToUsersNotOfContoso.csv

#All messages from senders of a domain.
Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-date).AddMonths(-1) -End (Get-Date) | Where-Object {$_.sender -like "*@Contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\ExchangeLogResults_SentFromContoso.csv

#All messages from senders not of a domain.
Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-date).AddMonths(-1) -End (Get-Date) | Where-Object {$_.sender -notlike "*@contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\ExchangeLogResults_NotSentFromContoso.csv
#endregion

#region Search by domain in EXO
#All messages with recipients like a domain.
Get-MessageTrace -Start (Get-date).AddDays(-30) -End (Get-Date) | Where-Object {$_.recipients -like "*@contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv \\someserver\someshare\ExchangeLogResults_ToUsersOfContoso$(get-date -f dd-MM-yyyy)services.csv

#All messages with recipients not like a domain.
Get-MessageTrace -Start (Get-date).AddDays(-30) -End (Get-Date) | Where-Object {$_.recipients -notlike "*@contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\temp\ExchangeLogResults_ToUsersNotOfContoso.csv

#All messages from senders of a domain.
Get-MessageTrace -Start (Get-date).AddDays(-30) -End (Get-Date) | Where-Object {$_.sender -like "*@Contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\ExchangeLogResults_SentFromContoso.csv

#All messages from senders not of a domain.
Get-MessageTrace -Start (Get-date).AddDays(-30) -End (Get-Date) | Where-Object {$_.sender -notlike "*@contoso.com"} | Select-Object Timestamp,SourceContext,Source,EventId,MessageSubject,Sender,{$_.Recipients} | Export-Csv C:\ExchangeLogResults_NotSentFromContoso.csv
#endregion

#region retrieve previous days logs

Get-MessageTrace -StartDate ([DateTime]::Today.AddDays(-1)) `
-EndDate ([DateTime]::Today) | Select MessageID,Received,*Address,*IP,Subject,Status,Size | `
Export-Csv "$((get-date ([DateTime]::Today.AddDays(-1)) -Format yyyyMMdd)).csv" -NoTypeInformation 

#endregion

#region spam/phishing info
Start-Process https://www.social-engineer.org/

Start-Process https://cofense.com/ #Phish Me solution
#endregion

#endregion

#region Users in O365 and on premises

#remove user
Remove-MsolUser -UserPrincipalName bandit@mcdeo.onmicrosoft.com

# Retrieve a list of all deleted users:
Get-MsolUser –ReturnDeletedUsers

# To restore all deleted users:
Get-MsolUser –ReturnDeletedUsers | Restore-MsolUser

# To restore a single deleted user:
Restore-MsolUser –UserPrincipalName  bandit@contoso.com

#Steps to look for and return deleted user
Get-Mailbox bandit

Get-MsolUser -ReturnDeletedUsers

Restore-MsolUser -UserPrincipalName User@contoso.com

Get-Mailbox User

#endregion

#region Merge one mailbox into another mailbox to recover from a deleted user.

# List the Soft Deleted Mailboxs and pick the one that needs to be imported 
$DeletedMailbox = Get-Mailbox -SoftDeletedMailbox | Select DisplayName,ExchangeGuid,PrimarySmtpAddress,ArchiveStatus,DistinguishedName | Out-GridView -Title "Select Mailbox and GUID" -PassThru

# Get Target Mailbox 
$MergeMailboxTo = Get-Mailbox | Select Name,PrimarySmtpAddress,DistinguishedName | Out-GridView -Title "Select the mailbox to merge the deleted mailbox to" -PassThru

# Run the Merge Command 
New-MailboxRestoreRequest -SourceMailbox $DeletedMailbox.DistinguishedName -TargetMailbox $MergeMailboxTo.PrimarySmtpAddress -AllowLegacyDNMismatch

# View the progress 
#Grab the restore ID for the one you want progress on. 
$RestoreProgress = Get-MailboxRestoreRequest | Select Name,TargetMailbox,Status,RequestGuid | Out-GridView -Title "Restore Request List" -PassThru

# Get the progress in Percent complete 
Get-MailboxRestoreRequestStatistics -Identity $RestoreProgress.RequestGuid | Select Name,StatusDetail,TargetAlias,PercentComplete

#Pass thru option in Out-Gridview demo
Get-Service | Out-GridView -PassThru

Get-Service | Out-GridView -PassThru > c:\temp\services.txt
notepad.exe C:\Temp\services.txt

Get-Mailbox | Out-GridView -PassThru > c:\temp\mailboxesToMerge.txt
#endregion

#region on-premises Exchange server logon functions and code for .ps1 files to run

#Specific server in function. Engineers can change the function and computer name per server in their org.
Function Connect-CON-EX2016N1 {
    $Credential = Get-Credential 
    $Session = New-PSSession -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ConnectionUri http://con-ex2016n1/PowerShell/ -Credential $UserCredential 
    Import-PSSession $Session
}

#Parameter requiring specific server
Function Connect-ExServer {   
    param ($Computer)
        $UserCredential = Get-Credential 
        $Session = New-PSSession `
            -Authentication Kerberos `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri http://$Computer/PowerShell/ `
            -Credential $UserCredential 
        Import-PSSession $Session -Prefix ONPrem
}

#Add snap-in for accessing on an exchange server. Use in scripts to access Exchange via normal PS.

#2007
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin;

#2010
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;

#2013/2016
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#endregion

#region Change Exchange Activation Preference for DB's on a server
function Set-MailboxActivationPreferenceSingleServer {
<#
.Synopsis
   Sets activation preference for databases on a single, specific server.
.DESCRIPTION
   A simple piping of cmdlets does not yield the proper format for the cmdlets to run.
   Have to gather the databases all of a specific server, then take just the name value, concatenate that with the server name, 
   then take that new value and pass it as the identity of the database to be updated with the activcation preference. 
.EXAMPLE
   Set-MailboxActivationPreferenceSingleServer -ServerName EXNode1 -ActivationPreference 1

   Sets the activation preference of server EXNode1 to a value of 1
.EXAMPLE
   Set-MailboxActivationPreferenceSingleServer -ServerName EX2013N2 -ActivationPreference 2

   Sets the activation preference of server EX2013N2 to a value of 2
.NOTES
   Created by Mike O'Neill to help customers migrate data centers and change activation preference for specific servers.
#>
param ($ServerName, $ActivationPreference = '1')
    $DBs = @(Get-MailboxDatabase -Server $ServerName).name #Obtains all Databases on a specific server
    $DBValues = foreach ($DB in $DBs) {$DB + "\$ServerName"} #Takes the name of each database and formats it correctly for the Identity value on the next line.
    foreach ($DBValue in $DBValues) {Set-MailboxDatabaseCopy -Identity $DBValue -ActivationPreference $ActivationPreference} #Loops through all DB's on the server and sets Activation preference.
}
#endregion

#region Hybrid mailbox moves

# Creating a list of migration batches with one user per batch
$Users = Get-Content c:\temp\users.txt | ForEach-Object {New-MigrationBatch -Name $_ -SourceEndpoint RemoteEndpoint1 -TargetDeliveryDomain cloud.contoso.com -Users $_ -NotificationEmails someone@contoso.com -AutoStart -AutoComplete}

#restart MRSProxy
$Exchange = Get-Content c:\temp\ExchangeServersInDataCenter.txt #Not using EMS

$Exchange = Get-ExchangeServer #Using EMS

Invoke-Command -ComputerName $Exchange | Get-Service MRSProxy | Restart-Service #Use this to restart MRSProxy once a list of Exchange servers is obtained

Invoke-Command -ComputerName $Exchange -ScriptBlock {Test-NetConnection -ComputerName  "smtp.office265.com" -Port 25 -InformationLevel "Detailed"} #Checks outbound TCP port 25


#Workflow function so all of the Exchange servers can have IIS restarted quickly
Function Restart-ExchangeIIS {
Get-content -path c:\temp\listOfServersToBounce.txt = $ServersINeed
$ExchangeServers = Get-ExchangeServer #Must be logged in remotely to Exchange server or running this via Exchange Management Shell
    Workflow Restart-AllExchangeIIS {
        Foreach -parallel ($ExchangeServer in $ExchangeServers) {            
                    Invoke-Command -ComputerName $ExchangeServer | Restart-Service w3svc -force 
                 }
     }
}

#restart IIS AutoD app pool, less abrasive than restarting IIS on all Exchange servers. 
Function Restart-AutoDAppPool {
<#
.Synopsis
   Restarts IIS AutoDiscover app pool on all Exchange servers in Organization
.DESCRIPTION
   Obtains all Exchange servers in the organization and 
   restarts the IIS AutoDiscover app pool
.EXAMPLE
   Restart-AutoDAppPool
.EXAMPLE
   Another example of how to use this cmdlet
.NOTES
   This function allows repeatable use of restarting the IIS app pool for AutoDiscover on Exchange servers.
   This function will not restart the entire IIS service, just the specific AutoD app pool.
#>
[cmdletbinding()]
Param()

    Get-ExchangeServer | foreach { 
            Write-Output "Recycling $_" ; Invoke-Command `
            -ComputerName $_.name `
            -ScriptBlock { Get-ChildItem IIS:\AppPools | `
            Where-Object { $_.name -eq "*autod*" } | ` #comment of this line
            Restart-WebAppPool 
        } 
    }
} 

Send-MailMessage `
    -Body "lkjlkjlkjlkjlskjdflksjdlfkjsldkfjslkj" `
    -to mike@contoso.com `
    -From guy@contoso.com


#endregion

#region Tenant to tenant migration
Start-Process "https://support.office.com/en-us/article/How-to-migrate-mailboxes-from-one-Office-365-tenant-to-another-65af7d77-3e79-44d4-9173-04fd991358b7?ui=en-US&rs=en-US&ad=US"
#endregion

#region EOP

#Send Grid service
Start-Service https://sendgrid.com/use-cases/transactional-email/

#Azure send bulk e-mail
Start-Process https://docs.microsoft.com/en-us/azure/

Get-MailDetailTransportRuleReport #Shows 7 days of messages

#when rules were last used.
$TransportRules = Get-TransportRule -ResultSize Unlimited
foreach ($rule in $transportrules) { Get-MailDetailTransportRuleReport -TransportRule $rule.name -StartDate (Get-Date).AddDays(-180) -EndDate (Get-Date) | Sort -Property Date -Descending | Select -First 1 date,messageid,subject,transportrule}

Function Get-EmailSecuritySettings {
<#
.Synopsis
   Present the current Email Security Settings of an o365 tenant.
.DESCRIPTION
   Shows the: SPF, DKIM, and DMAC settings of a tenant. 
   This function confirms the status of these settings in order to review if the different e-mail security settings have been enabled and/or are configured.
.EXAMPLE
   Get-EmailSecuritySettings

   Gets the currently logged in tenant settings.
.EXAMPLE
   Get-EmailSecuritySettings -AcceptedDomains contoso.com

   Gets the contoso.com in tenant settings.
.NOTES
   Creator: Matt Fields - PFE
   Updated: Mike O'Neill - PFE

   Version 1.0 - created code and functioned up for repeatable usage. 
#>
[cmdletbinding()]
param($AcceptedDomains = $(Get-AcceptedDomain))

Write-Verbose "Accepted domains obtained: $($AcceptedDomains)"
ForEach ($Domain in $AcceptedDomains) {

#SPF Check
$SPF = Resolve-DnsName -Type TXT -Name $Domain -ErrorAction SilentlyContinue | Where-Object {$_.Strings -like "v=spf1*"}
Write-Verbose "Domain working on: $($domain)" 

If (!($SPF)) {
        Write-Host -ForegroundColor White -BackgroundColor Red "SPF has not been configured for $($Domain.DomainName)"
    }

Else {
        $SPF | ft -AutoSize
    }

#DKIM Check
$DKIM = Get-DKIMSigningConfig $Domain.DomainName

IF ($DKIM.Enabled -eq 'True') {
    $DKIM = Resolve-DnsName -ErrorAction SilentlyContinue -Type CNAME -Name "selector1._domainkey.$($Domain.DomainName)"

    IF (!($DKIM)) {
            Write-Host -ForegroundColor White -BackgroundColor Red "DKIM is enabled, but DNS entries have not been created for $($Domain.DomainName)"
        }

     }
Else {
        Write-Host -ForegroundColor White -BackgroundColor Red "DKIM is not enabled for $($Domain.DomainName)"
     }

#DMARC Check
    $DMARC = Resolve-DnsName -ErrorAction SilentlyContinue -Type TXT -Name "_dmarc.$($Domain.DomainName)" | Where-Object {$_.Strings -like "v=DMARC1*"}

If (!($DMARC)) {
            Write-Host -ForegroundColor White -BackgroundColor Red "DMARC has not been configured for $($Domain.DomainName)"
      }

Else {
         $DMARC | ft -AutoSize
     }

    } 
} 


#endregion

#region Auto-Archiving process
Start-Process https://support.office.com/en-US/article/Enable-unlimited-archiving-in-Office-365-e2a789f2-9962-4960-9fd4-a00aa063559e

#endregion

#region MFA remote PS
Start-Process 'https://technet.microsoft.com/en-us/library/mt775114(v=exchg.160).aspx'
#endregion

#region Misc

#Monitoring, reporting, and message tracing in Exchange Online
Start-Process "https://technet.microsoft.com/en-us/library/jj200725(v=exchg.150).aspx"

#Find expiring certificates
Cd cert:
Get-ChildItem cert:LocalMachine –recurse | where-object {$_.NotAfter –le (Get-Date).AddDays(-365) –And $_.NotAfter –gt (Get-Date).AddDays(990)} | Select thumbprint, subject, issuer

#get date options
(Get-Date)
(Get-Date).AddDays(-30)
get-date | Get-Member



(Get-Date).AddMonths(-10)

#disk speed test
Winsat disk -drive c -ran -write -count 10

psedit "ps1 file"
#endregion

#region updating calendaring in tenant

Get-MailboxFolderStatistics bandit | ? {$_.FolderPath -like "/Calendar/*"} | ft folderpath,folderid 

#endregion

#region MFA process with MRM
Start-Process https://gallery.technet.microsoft.com/Powershell-script-to-2489e63b

#Attributes that are being sync'd to cloud:
Start-Process https://docs.microsoft.com/en-us/azure/active-directory/connect/active-directory-aadconnectsync-attributes-synchronized

#endregion

#region Nullify user accounts that are broken by the mailbox move process not clearing out the user attribute

cd ad:\ #Using the AD provider
Set-Location 'DC=contoso,DC=com' #Domain name to map to
sl 'OU=UsersToFix' #Users OU to fix

#For msExchMailboxMoveStatus
Get-Item -Filter "msExchMailboxMoveStatus=*" -Path * #Check to see which mailboxes are impacted

Clear-ItemProperty -Filter "msExchMailboxMoveStatus" -Path * -Name msexchmailboxmovestatus -WhatIf

#For msExchMailboxMoveStatus
Get-Item -Filter "msExchMailboxMoveRemoteHostName=*" -Path * #Check to see which mailboxes are impacted

Clear-ItemProperty -Filter "msExchMailboxMoveRemoteHostName" -Path * -Name msExchMailboxMoveRemoteHostName -WhatIf

#For removing Leg DN value if Disable-Mailbox breaks
Get-Item -Filter "legacyExchangeDN=*" -Path * #Check to see which mailboxes are impacted

Clear-ItemProperty -Filter "legacyExchangeDN=*" -Path * -Name legacyExchangeDN -WhatIf #To check which mailboxes will be impacted
Clear-ItemProperty -Filter "legacyExchangeDN=*" -Path * -Name legacyExchangeDN #To clear impacted mailboxes

#endregion

#region Exchange security

Start-Process https://gallery.technet.microsoft.com/Update-Stale-Health-f77ad037

#endregion

#region Focus Inbox
Start-Process https://www.csssupportwiki.com/index.php/curated:Focused_Inbox

#endregion 

#endregion

#region Active Directory demos
Start-Process "http://social.technet.microsoft.com/wiki/contents/articles/23313.notify-active-directory-users-about-password-expiry-using-powershell.aspx"

$env:windir
$env:ProgramFiles
$DC = Get-ADDomainController; Restart-Computer $DC

#modify attributes in attribute editor
Start-Process https://blogs.technet.microsoft.com/heyscriptingguy/2013/03/21/use-the-powershell-ad-provider-to-modify-user-attributes/
Start-Process https://blogs.technet.microsoft.com/heyscriptingguy/2013/03/18/playing-with-the-ad-drive-for-fun-and-profit/

Function Restart-AllDCs {
    $DC = Get-ADDomainController
    Restart-Computer $DC
}

Function Restart-AllDCsQuickly {
$DCs = Get-ADcomputer
    Workflow Restart-AllDomainControllers {
        Foreach -parallel ($DC in $DCs) {            
                                            Stop-Computer $DC
                                        }
     }
}

Invoke-Command -ComputerName dc1,dc2 -ScriptBlock {Test-ComputerSecureChannel}
Invoke-Command -ComputerName $dc -ScriptBlock {Test-ComputerSecureChannel}

#run with multiple prefixes
$sessions = New-LabPSSession -ComputerName (Get-LabVM -Role ADDS)

foreach ($session in $sessions)
{
    Import-Module -Name ActiveDirectory -PSSession $session -Prefix $session.ComputerName
}

Get-xDC1ADDomain
Get-xDC2ADDomain 


#Find user password expiration times
Function Get-UserPSExpiration {
<#
.Synopsis
   Get the value of when a users' password will expire.
.DESCRIPTION
   Query AD DS and obtain the msDS-UserPasswordExpiryTimeComputed attribute. 
   This shows when a password for a user will expire.
   Use this function when auditing or preparing to make a change to the expiration of a user's password.
.EXAMPLE
   Get-UserPSExpiration
.NOTES
    Code taken from here, then compiled into a re-usable function
    Start-Process https://blogs.technet.microsoft.com/poshchap/2014/02/21/one-liner-get-a-list-of-ad-users-password-expiry-dates/
.FUNCTIONALITY
   Gets the expiration date/time of when a users' password will expire. 
#>
[cmdletbinding()]
Param()

    Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} `
    –Properties “DisplayName”, “msDS-UserPasswordExpiryTimeComputed” | `
    Select-Object -Property “Displayname”,@{Name=“ExpiryDate”;Expression={[datetime]::FromFileTime($_.“msDS-UserPasswordExpiryTimeComputed”)}}
}

#region How to find RIDs
function Get-RIDsRemaining {
    param ($domainDN)
    $de = [ADSI]”LDAP://CN=RID Manager$,CN=System,$domainDN”
    $return = new-object system.DirectoryServices.DirectorySearcher($de)
    $property= ($return.FindOne()).properties.ridavailablepool
    [int32]$totalSIDS = $($property) / ([math]::Pow(2,32))
    [int64]$temp64val = $totalSIDS * ([math]::Pow(2,32))
    [int32]$currentRIDPoolCount = $($property) – $temp64val
    $ridsremaining = $totalSIDS – $currentRIDPoolCount
    Write-Host “RIDs issued: $currentRIDPoolCount”
    Write-Host “RIDs remaining: $ridsremaining”
}

#endregion

#endregion

#region Azure AD demos
Start-Process https://docs.microsoft.com/en-us/powershell/azuread/v2/azureactivedirectory?redirectedfrom=msdn #Azure AD v2.0 module

#Convert Immutable ID to object GUID and back
start-process https://gallery.technet.microsoft.com/office/Covert-DirSyncMS-Online-5f3563b1

#Azure Identity troubleshooting tools
Start-Process https://adfshelp.azurewebsites.net/
#endregion

#region Skype for Business/Lync 

New-MailboxSearch -Name 'blabla' -MessageTypes @('im') -SourceMailboxes @('c7f83c1f-4e45-4c28-a80f-b57b7441490f') -EstimateOnly:$true 
#endregion

#region PowerShell for IT admin Part 1

#region Version info
# ps version not in ISE
$PSVersionTable.PSVersion
Powershell -version 3

Get-Clipboard

#endregion

#region Module 1 Introduction
Get-Service

Get-Service -ComputerName win-8

#endregion Module 1

#region Module 2 Commands 1
Get-Service -Name Spooler
Start-Service -Name Spooler -Force

Get-Process
Restart-Service -Name WSearch -Verbose -ErrorAction SilentlyContinue

get-command -Name new-Mailbox -Syntax

Get-Process -Name Netlogon -ErrorAction SilentlyContinue
Get-Process Netlogon -ea 0

Send-MailMessage -Body "This is the body of text" `
    -Attachments file.txt `
    -Bcc someone@contoso.com `
    -From someoneelse@contoso.com 

#What if process
Stop-Process -Name * -WhatIf
Stop-Process -Name notepad -Verbose

Start-Process -name notepad #only works as admin
Get-Command -Name Start-Service -Syntax

Get-Command | more

Restart-Service -Name Netlogon -WhatIf

Clear-Host

get-process netlogon -ea silentlycontinue
Get-process logon

Get-Service BITS ; Get-Process System

Get-Service BITS
Get-Process System
#Show command, leverage the gui in ISE to be able to import command
Get-Command
Show-Command

Get-Command -Name *user*
Get-Command -Verb get
Get-Process -name
Get-Process -ComputerName localhost -Debug
Get-Process
Get-Mailbox -Archive -SoftDeletedMailbox

Get-Help Get-ChildItem -examples
get-help Get-Mailbox -Examples

Get-Service -Name s* | Where-Object {$_.status -eq "stopped"}

Get-Alias | fw -Column 4
New-Alias -Name list -Value Get-ChildItem

Remove-Item alias:list

Get-Mailbox bandit
Get-Mailbox -Identity bandit

#'Get' is assumed
service
mailbox
date

#endregion Module 2

#region Module 3 Pipeline 1
Get-Service -Name w32time , BITS | Stop-Service #Run as Admin

#domain info for logged on user
whoami.exe

whoami.exe | Split-Path -Parent 

whoami.exe | Split-Path -Leaf 

#View process and output on screen
Get-Process | Sort-Object ws | Select-Object -last 5

Get-EventLog -LogName Application | Group-Object EntryType 

Get-ChildItem C:\Temp | Measure-Object -Property IsReadOnly -Sum

#FL
Get-Process -Name powershell | Format-List *

Get-Process -Name powershell | Format-List -Property Name, BasePriority, PriorityClass 

#FT
Get-Process | Format-Table –Property name,workingset,handles 

Get-Process | Format-Table -Property Name,Path,WorkingSet -AutoSize -Wrap

Get-ChildItem | Format-Wide -Column 3 

Get-Alias | Format-Wide -AutoSize

Get-Service | Measure-Object






#Exporting discussion
Get-Service | Export-Csv "c:\temp\services.csv" -NoTypeInformation
notepad.exe C:\temp\services.csv
Get-Service | Export-Csv "c:\temp\$(get-date -f dd-MM-yyyy)services.csv" -NoTypeInformation

Get-Service | Out-GridView -PassThru


Get-Mailbox | Out-GridView -PassThru | Set-Mailbox -ProhibitSendReceiveQuota 0k
Get-Mailbox | Out-GridView -PassThru | Out-File C:\Temp\mailboxesIReallyWant.txt

Get-Service | Out-GridView -PassThru | Out-File c:\temp\services_Selected.txt 

notepad.exe C:\temp\services_selected.txt 

$groupedEvents = Get-EventLog -LogName application | Group-Object EntryType 
$groupedEvents  | Format-Table -AutoSize 
$groupedEvents = $null

#Lab hints
Get-Command -CommandType Cmdlet #tab through the options
Get-Command -CommandType Function | Measure-Object

Get-ChildItem 'C:\Program Files\Managed Defender' #This is a remark or a comment
Get-ChildItem C:\Windows\System32 -file | Measure-Object -Property Length -Sum

Get-ChildItem C:\Windows\System32 -file | Group-Object -Property extension | Sort-Object -Property Count -descending | Select-Object –first 5


#Labs 3.4.2
#6:
Get-Service | Export-Csv -Path 'C:\PShell\Labs\Lab_3\services.csv'

#7: 
Import-Csv –Path 'C:\PShell\Labs\Lab_3\services.csv' | Group-Object –Property ServiceType

#8: 
Get-Process | Export-csv -Path 'C:\temp\processes.csv'

#9: 
Import-Csv -Path 'C:\Temp\processes.csv' | Get-Process | Sort-Object -Property peakworkingset -Descending | Format-Table name,handles,cpu,peakworkingset

Import-Csv -Path 'C:\Temp\processes.csv' | Select-Object name,handles,cpu,@{Name="PeakWorkingSet";Expression={[Int]$_.PeakWorkingSet}} | Sort-Object -Property peakworkingset -Descending | Format-Table

#3.4.2 Text sorting and how to fix it for integer sorting



#endregion Module 3

#region Module 4 Commands 2

Function Get-ServiceInfo 
{
    Get-Service -Name Spooler -RequiredServices -ComputerName Localhost
}

Function Get-ServiceInfoVariable 
{
Param ($svc, $computer)
    Get-Service -Name $svc -RequiredServices -ComputerName $computer
}

Get-Service | Where-Object {$_.Status -eq "running"}

$env:COMPUTERNAME

TCP port 5985 #Default PS remote port
TCP port 5986 #Default SSL PS remote port

#Labs
$sb = {Get-WinEvent –LogName System –MaxEvents 5}
# & $sb

$ab = {
Param ($parameter1)
<statement list>
}

$ab = {param ($featureName) Get-WindowsOptionalFeature -Online -FeatureName $featureName}

#  &$ab -featurename windowsmediaplayer

Function Get-Feature 
    {
        param ($featureName) Get-WindowsOptionalFeature -Online -FeatureName $featureName
    }

#endregion Module 4

#region Module 5 Scripts

Get-ExecutionPolicy –List

Ping localhost

Test-Connection localhost

New-Alias -Name ping -Value Test-Connection

Remove-Item alias:ping

Get-Process system

Function Get-Process {"This is not the get-process built-in function."}

Get-Process

Get-Command Get-Process -All #Shows all of the 'get-process' values, not just the current precedent one

Microsoft.PowerShell.Management\Get-Process -Name System
Get-Process
Set-Alias Get-Process Get-Service
Get-Process 

Get-Command Get-Process -All

Remove-Item function:get-process

Function Ping {
    Ping.exe localhost
    Start-Process http://www.google.com
}
Set-ExecutionPolicy

Function Get-Process {Microsoft.PowerShell.Management\Get-Process -Name System}

#Kill is alias for ‘stop-process’. 
Get-process *ise* | kill

#For help with the labs to ensure paths
Test-Path C:\temp

#endregion Module 5

#region Module 6 Help
Get-Help -ShowWindow 

Get-Help Get-Service -ShowWindow

Function Get-SysLogNN {
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Get-SysLogNN -LogName Application -Newest 10
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>

param ($LogName="application",$NumberOfEvents)
    Get-EventLog -LogName $LogName -Newest $NumberOfEvents
}

Function Get-SyslogNN {

 <#
.Synopsis
   This function is something that works. 
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet example is cool
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
param ($eventlogChannel="application",$Number="10")
    Get-EventLog -LogName $EventLogChannel -Newest $Number
}

#endregion Module 6

#region Module 7 Object Modules

Get-User | Get-Member

Get-Mailbox | Get-Member

Get-ChildItem c:\windows\windowsupdate.log | Get-Member 

$csvhash = Get-FileHash c:\windows\windowsupdate.log
$csvhash | Get-Member
$csvhash.Path
$csvhash.Algorithm
 

Get-Mailbox | Get-Member 
Get-Mailbox bandit | fl


$mbx = Get-Mailbox bandit
$mbx.UMenabled

$allmbx=Get-Mailbox
$allmbx.

(Get-FileHash c:\windows\windowsupdate.log) | Get-Member 
(Get-FileHash c:\windows\windowsupdate.log).hash
(Get-FileHash c:\windows\windowsupdate.log).path
$date=get-date
(get-date).AddTicks(-$date.Ticks)

Get-Mailbox | Get-Member 

Get-Mailbox | Get-Member -MemberType Method

Get-Process -Name spoolsv | Get-Member -MemberType Method
Get-Process spoolsv | Get-Member -MemberType Properties
Remove-Item function:get-process

Get-WmiObject win32_bios | Get-Member -membertype properties
(Get-WmiObject win32_bios).status
Get-ExchangeServer | Format-Table (Get-WmiObject win32_bios).Status

Get-Date | Get-Member 

(Get-Date).addmonths(-5) #monthly report
get-date | fl *
(Get-Date).GetType()

Get-Process | Get-Member

Get-Item C:\temp\PowerShell_transcript.txt | gm
$File = Get-Item C:\temp\PowerShell_transcript.txt
$file.LastAccessTime
$file.LastAccessTime = (get-date).AddYears(-200)

#Additional time stamp information
Start-Process https://blogs.technet.microsoft.com/heyscriptingguy/2012/06/01/use-powershell-to-modify-file-access-time-stamps/

#  & 'C:\Program Files\Common Files\microsoft shared'

$date=Get-Date -f dd-MM-yyyy-HH-mm

New-Item "c:\temp\$env:computername from dev team $date.txt"

Write-Host
Write-Host "This is todays' date $date and it is cool!" -ForegroundColor Black -BackgroundColor White

#endregion Module 7

#region Module 8 Operators 1

Get-Help about_Regular_Expressions

'Pear' -clike '*contoso*'


$service = Get-Service bits
$service.Status -eq 'Running'

(Get-Process).Name -contains 'Notepad' 

-not (Test-Path C:\Windows) 

$NumberA = 4
$NumberB = 10

(($NumberA -lt 8) -and ($NumberA -gt 'a')) -and ($NumberB -le 1)


((get-windowsoptionalfeature -online|select-object -expandproperty Featurename) -like "*Play*") -xor ((get-windowsoptionalfeature -online|select-object -expandproperty Featurename) -like "*Media*")

Get-ChildItem C:\Windows\System32\[a-d][xcp][a-l]*

2,3,4,5,6 -lt 5

#Numeric of $True/$False
Start-Process https://blogs.msdn.microsoft.com/powershell/2006/12/24/boolean-values-and-operators/

function test ($VALUE) {
if ($VALUE) {
    Write-Host -ForegroundColor GREEN “TRUE”
    } 
else {
    Write-Host -ForegroundColor RED   “FALSE”
    }
}

#endregion Module 8

#region Module 9 Pipeline 2

Get-Mailbox | ft name,alias

Get-Process | Get-Member

Get-Process | Where-Object {$psitem.workingset64 -gt 100MB}
Get-Process | ? {$_.ws -gt 100MB} 
Get-Process -pipelinevariable CurrentProcess | Where {$CurrentProcess.ws -gt 100Mb} 

Get-Process | Get-Member

Get-Service | Get-Member

Get-Service net* | ForEach-Object {"Hello " + $PSItem.ServiceName} 
Get-Service net* | ForEach {"Hello " + $_.Name + " how are you today?"} 
Get-Service net* | ForEach {"Hello $($_.Name) how are you today?"}
Get-Service net* | % {"Hello $($_.name)"} 
Get-Service net* | % {"Hello $_"} 

(Get-EventLog -LogName Application).TimeWritten.DayOfWeek

Get-Service | Where-Object {$_.Status -eq "running"}
Get-Service | Where Status -eq Running 
Get-Service | ? Status -EQ running 

Get-Process -PipelineVariable CurrentProcess | Where-Object {$CurrentProcess.ws -gt 100MB}

# What is ws? How to get it?

Get-Service net* | ForEach-Object {"Hello  $($_.Name)"} 
Get-Service net* | % {"Hello " + $_.Name} 

Get-Process | ForEach-Object ID
Get-Process | Get-Member

(Get-Process).id
Get-Process | Get-Member
(Get-Process).pagedmemorysize
(Get-Process).processname
Get-Process | % processname
Get-Process | Where-Object {$_.processname} | fl name


Get-Eventlog system | Get-Member

(Get-EventLog –Log System).TimeWritten.DayOfWeek | Group-Object

Get-Date | Get-Member

#demo get-mailbox for get-date info

Get-EventLog -LogName System -Newest 5 | ForEach-Object -Begin {Remove-Item c:\temp\Events.txt; Write-Host "Start" -ForegroundColor Yellow} -Process {$_.Message | Out-File -Filepath c:\temp\Events.txt -Append} -End  {Write-Host "End" -ForegroundColor Green; notepad.exe c:\temp\Events.txt} 

Get-EventLog -LogName System -Newest 5 | ForEach-Object `
 -Begin {Remove-Item c:\temp\Events.txt -ErrorAction SilentlyContinue; Write-Host "Start" -ForegroundColor Yellow} `
 -Process {$_.Message | Out-File -Filepath c:\temp\Events.txt -Append} `
 -End {Write-Host "End" -ForegroundColor Green; notepad.exe c:\temp\Events.txt} 


Get-EventLog -LogName system -Newest 20 | Create-FileOfEvents 

 function Create-FileOfEvents {
    Begin
    {
        Remove-Item c:\temp\Events.txt
        Write-Host "Start" -ForegroundColor Yellow
    } 
    Process
    {
         $_.Message | Out-File -Filepath c:\temp\Events.txt -Append
    } 
    End
    {
        Write-Host "Process is now complete" -ForegroundColor Green
        notepad.exe c:\temp\Events.txt
    } 
}  



Get-Help Restart-Computer -Parameter ComputerName

Get-Help Restart-Computer -full

1..10 | % {Send-MailMessage -To User@contoso.com -From external@external.com -SmtpServer ExchangeServerNameHere -Subject "Test Message $_" -Body "This is the body of Message $_" ; write-host “Sending Message $_”}

1..3600 | ForEach-Object {Send-MailMessage -To User@ReceivingTenant.com -From external@SendingTenant.com -SmtpServer smtp.office365.com -Subject "Test Message $_" -Body "This is the body of Message $_" ; write-host “Sending Message $_”}

1..1000000000 | % {New-Mailbox -Name User$_ -Alias User$_ –DisplayName `
“User$_” -Password (ConvertTo-SecureString "Password1" -AsPlainText -Force) `
-UserPrincipalName "user$_@contoso.com" -OrganizationalUnit “contoso.com/Accounts"}

#endregion Module 9

#region Module 10 Providers

Get-PSProvider

Get-PSDrive

New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT

Get-Content Env:\ProgramFiles ; Get-Content Env:\COMPUTERNAME

New-PSDrive -Name H -PSProvider filesystem -Root \\someserver\someshare
Get-PSDrive
Remove-PSDrive -Name HKCR
cd cert:\

Get-Item | Get-Member
Get-Item C:\Windows | Get-Member
(Get-Item C:\Windows).lastaccesstime

Set-Location -Path $env:ProgramFiles
Set-Location -Path c:\

Test-Path $env:windir

Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\O365_Logon"

$create = Join-Path -path c: -ChildPath temp2
md $create

#show -relative option
Resolve-Path c:\prog* -Relative 

Remove-Item -Path hkcu:\

$env:OS

#Load AD provider
Del AD:\* -Recurse

#endregion Module 10

#region Module 11 Variables & data types

Get-Help about_Automatic_Variables

$a = 123
$a 

$b = 'As easy as $a'
$e = 'as easy as '+ $a
$e
$b 

$c = "As easy as $a"
$c

$d = "This is a line of text."
$d

$lString = '
As
easy
as
$a
' 

$eString = "
As
easy
as
$a
" 
$eString
$lString


#Why Here String?

$X = @"
"Curiouser and curiouser!" cried Alice (she was so much surprised, 
that for the moment she quite forgot how to speak good English); 
"now I'm opening out like the largest telescope that ever was! 
Good-bye, feet!" 
"@

#Variable sub-expression demo
$a = $null
$a = Get-Service -Name ALG 
$a.GetType().FullName
Write-Host "service: $a"
Write-Host "service: $a.name" 
Write-Host "service: $($a.displayname)" 

Write-Host 'service:' $a 'that I asked for.'
Write-Host $a

(1024).GetType().FullName 

(1.6).GetType().FullName
 
(1tb).GetType().FullName 

$mynotnumber = "000123"

$MyNumber = [int]"000123" 
$MyNumber 

$MyNumber.GetType().FullName
$Mynotnumber.GetType().FullName

$MyNumberAlso = "000123"
[int]$MyNumberAlso
$MyNumberAlso.GetType().FullName 

$test = [int]"abc123"
$test

$MyNumberAlso.GetType().FullName

#undo typecast
[object] $MyNumberAlso

[char] | Get-Member -Static | Measure-Object
[char] | Get-Member |Measure-Object

[char]::IsWhiteSpace(" ")


[math] | gm
[math] | gm -Static

[math]::pow(2,3) 
[math]::pi 
[math]::round([math]::pi,16)
[math]::Round(12.5) #Rounding to nearest 'even' number. Even numbers are better, as the others are just 'odd'.
[math]::Round(12.5, 0, "awayfromzero")
[math]::Round(13.5, 0, "awayfromzero")
[int]12.5
[int]13.5

33..255 | ForEach-Object {
    Write-Host "Decimal: $_ = Character: $([Char]$_)"
    Start-Sleep -Seconds 1
}

write-host "The date is: (get-date)"
write-host "The date is: $(get-date)"
write-host "The date is: "(get-date)

"27/12/2013" -as [datetime]

$cmd = "Get-Process"
#  & $cmd
$cmdNew = Get-Process

$a = 123
Write-Host "`$a is $a"

Write-Host "There are four line breaks`n`n`n`nhere. "

#More about types:

$a = (Get-Date).DayOfWeek
$b = Get-Date | Select-Object DayOfWeek

$a;$b

$a.GetType();
$b.GetType();

$b.DayOfWeek -eq $a

#Static Classes and Methods
Start-Process https://msdn.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/using-static-classes-and-methods

#endregion Module 11

#region Module 12 Operators 2

-split "1 a b" 

"1, a b" -split "," 

"Windows PowerShell 4.0" –replace "4.0","5.0"

"Windows PowerShell 4.0" –ireplace "win",""


$MyArray = 'Smith','John',123.456789

“Custom Text" -f $MyArray

"First name is: {1},  Last name is: {0}" -f $MyArray

"Using a Format Specifier {2:N1}" -f $MyArray

#More date value options
(Get-Date).AddDays(-7).ToString('MM-dd-yyyy')
 
'{0:dd}-{0:MM}-{0:yyyy}' -f (Get-Date).AddDays(-7)

#region characters in PS CSV output issue
# fix gibberish text logs

# file to clean
param ($file = (Read-Host "File"))

# raw file contents in bytes
$a = Get-Content $file -Encoding Byte -ReadCount 1 | %{"{0:X2}" -f $_}

# remove any instances of fffe, the characters that seem to cause the formatting to break.
$b = ($a -join "") -replace 'fffe',''

# split the bytes back out into 2 digit groups and convert back to ANSI
$c = @()
for ($i = 0; $i -lt $b.Length; $i += 2) {
    $c += [byte]("0x$($b[$i])$($b[$i + 1])")
}

# convert to ANSI and write to file
$path = Split-Path $file -Parent
$name = (Split-Path $file -Leaf) -replace ".LOG","_fixed.LOG"

[System.Text.Encoding]::ASCII.GetString($c) | Out-File "$path\$name" -Encoding ascii 

#endregion

#endregion Module 12

#region Module 13 Arrays

$processarray = Get-Process
$processarray

$processarray | Select-Object -last 5
$processarray[0..5]

$array = 22,5,10,8,12,9,8
$array[0..308]
$array[-1]
$array[-1..-5]
$array[0,-1]

$b = @(22,33,44)
$b

$array.Count 
$processarray.length

$array2 += 999
$array += "Scottie"
[string]$array += 123,456,98709870987

$array2 | Sort-Object -Descending

[array]::Sort($array2); $array2[0,-1] 

$array | Get-Member 
Get-Member -InputObject $array 

$array = $null

#Get-Unique
$new = 1,2,3,4,5,6
$old = 1,2,5,7,8,9,10

$allexe = $new + $old
$allexe | Sort-Object | Get-Unique -OutVariable AllExeUnique
$AllExeUnique

#endregion Module 13

#region Module 14 Hash Tables

#Hash Table
$Server = @{'HV-SRV-1'='192.168.1.1' ; Memory=64GB ; Serial='THX1138'}   
$Server

#Hash table using here string
$string = @"
Msg1 = Hello
Msg2 = Enter an email alias
Msg3 = Enter an username
Msg4 = Enter a domain name
"@ 

$string

ConvertFrom-StringData -StringData $string

#Already know how to look up running tasks
Get-Service | Where-Object {$PSItem.Status -eq "Stopped"} | Measure-Object
Get-Service | ? {$_.Status -eq "Stopped"} | measure

#Another way
$svcshash = Get-Service | Group-Object status -AsHashTable -AsString 

$svcshash 
$svcshash.Count
$svcshash.Values
$svcshash.Stopped | Measure-Object
$svcshash.Running | measure

Get-EventLog -LogName Application -Newest 15 -EntryType Warning -ComputerName localhost

$params = @{
  LogName      = "application"
  Newest       = 15
  EntryType    = "Warning"
  ComputerName = "localhost"
}

Get-EventLog @Params 


#Send message to yourself or DL when events appear on a server
$30MinutesAgo = [DateTime]::Now.AddMinutes(-30)
    $messageParameters = @{ 
        Subject = "User Account Locked" 
        Body = Get-EventLog "Security" | Where {$30MinutesAgo -le $_.TimeWritten -and $_.eventid -eq 4740} | Format-List | Out-String 
        From = "PDC-Server@contoso.com" 
        To = "User@contoso.com", "DL@contoso.com" 
        SmtpServer = "Exchange-Server.contoso.com" 
    } 

Send-MailMessage @messageParameters

Send-MailMessage -Subject "User account" -Body Get-EventLog -From someone@contoso.com -to 

#Speed comparison: 
$params = @{
    LogName = "system"
    Newest = 15
    EntryType = "Warning"
    ComputerName = "localhost"
}
Measure-Command {Get-EventLog @Params} 
Measure-Command {Get-EventLog `
-LogName System `
-Newest 15 `
-EntryType Warning -ComputerName localhost} 

#region Another splat example
function Get-EventLogAppCrashes {
    [CmdletBinding()]
    [OutputType()]

    param ()

    process
    {
        $eIds = [Int64[]] 1000,1023,1008
        $logName = 'Application'
        $afterDate = (Get-Date).AddDays(-14)

        $splatParams = @{
            LogName = 'Application';
            InstanceId = 1000,1023,1008;
            EntryType = 'Error'
            After = $afterDate
        }

        $events = (Get-EventLog @splatParams | Select -Property MachineName,TimeGenerated,Message,Source,EventID,EntryType,Category,ReplacementStrings)
        
        Write-Output ($events | Sort MachineName,TimeGenerated -Descending)
    }
}

Get-EventLogAppCrashes | Export-Csv C:\Temp\Splat.csv -NoTypeInformation


#endregion

#endregion Module 14

#region Module 15 Flow Control
#Logic issues:
<# becareful of logic

ME:"Please go to the store and buy a carton of milk and if they have eggs, get six." 

Roommate came back with 6 cartons of milk.

ME:"Why did you buy six cartons of milk? I only need 1!" 

Roommate: "But you said if they had eggs get 6, they had eggs!" 

#>

$Services = Get-Service
'There are a total of ' + $Services.Count + ' services.'
ForEach ($Service in $Services)
{
    $Service.Name + ' is ' + $Service.Status
}



Switch (Get-ChildItem -Path c:\)
{
    "program*" {Write-Host $_ -ForegroundColor Green}
    "windows" {Write-Host $_ -ForegroundColor Cyan}
} 

switch –Wildcard (Get-ChildItem -Path c:\)
{
    "program*" {Write-Host $_ -ForegroundColor Green}
    "windows"  {Write-Host $_ -ForegroundColor Cyan}
} 

$c = 0
While ($c -lt 4)
{
    $c++ 
    if ($c -eq 2) {Continue}
    Write-Host $c
}  


 function Test-Return ($val)
{
    if ($val -ge 5) {return $val}
    Write-Host "Reached end of function"
}

#Exchange large array issue: 

#region working with large data sets
#This array will hold just the properties we need for each mailbox. This helps to conserve memory
$mailboxes = @()

#Run the Get-Mailbox command and store the properties we need in $mailboxes
Get-ExchangeServer EXSrv* | Get-Mailbox | ForEach-Object{
       $mailbox = New-Object PSObject

       $mailbox | Add-Member NoteProperty Name $_.Name
       $mailbox | Add-Member NoteProperty WindowsEmailAddress $_.WindowsEmailAddress
       $mailbox | Add-Member NoteProperty OrganizationalUnit $_.OrganizationalUnit

       $mailboxes += $mailbox
}

#Loop through and process each mailbox
foreach ($mailbox in $mailboxes)
{
    ...
    Write-Host "Mailbox Name: $($_.Name)"
    Write-Host "Email: $($_.WindowsEmailAddress)"
    Write-Host "OU: $($_.OrganizationalUnit)"
    ...
} 

$mailboxes = Get-ExchangeServer ExchSrv* | Get-Mailbox

foreach ($mailbox in $mailboxes)
{
    ...
    Write-Host "Mailbox Name: $($mailbox.Name)"
    ...
}
#endregion

#endregion Module 15

#region Module 16 Scope

$profile | Format-List -Property * -Force 
$profile

#Profile info
Test-Path $profile

New-Item -path $Profile -type file –force

notepad $profile

get-themes

#functions
Function Get-ServerData
{
    Param ($ComputerName)
    Get-CimInstance win32_OperatingSystem -ComputerName $ComputerName
}


Get-ServerData

#endregion Module 16

#region Module 17 Modules

Get-Module -ListAvailable 

$env:PSModulePath -split ';'

Get-Command -Module o365_logon

Get-Command -Module foomod

Import-Module foomod -Force

New-ModuleManifest -path C:\Windows\System32\WindowsPowerShell\v1.0\Modules\FooMod\foomod.psd1 #edit cmdlets '*'

#Force module discovery if in wrong path
Start-Process http://www.energizedtech.com/2016/08/making-the-configuration-manager-powershell-module-discoverable-two-lines-2.html

#endregion Module 17

#endregion

#region PowerShell for IT admin Part 2

#region Module 1 Review of Part 1 course

#endregion

#region Module 2 Remoting

#region $Using
#local Variable
$ServiceName = 'Bits'

#ServiceName not in remote runspace
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {"ServiceNames: " + $ServiceName}
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {$ServiceName | Set-Service}

#PSv2 - Argument and incoming parameter
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {Param($ServiceName) ; "ServiceNames: " + $ServiceName} -ArgumentList $ServiceName
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {Param($ServiceName) ; $ServiceName | Set-Service} -ArgumentList $ServiceName

#PSv3 - Using Prefix
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {"ServiceNames: " + $Using:ServiceName}
Invoke-Command -ComputerName 2012R2-MS -ScriptBlock {$Using:ServiceName | Set-Service}

TCP port 5985 #Default PS remote port
TCP port 5986 #Default SSL PS remote port

#Control CIM session protocol:
$CimSessiosOption = New-CimSessionOption -Protocol 
$CIMSession = New-CimSession -ComputerName $computer -SessionOption $CimSessiosOption 
Get-CimInstance -ClassName $class -CimSession $session  

#endregion

#multiple prefixes
$sessions = New-LabPSSession -ComputerName (Get-LabVM -Role ADDS)

foreach ($session in $sessions)
{
    Import-Module -Name ActiveDirectory -PSSession $session -Prefix $session.ComputerName
}

Get-xDC1ADDomain
Get-xDC2ADDomain 


#Windows Firewall: Windows Remote Management (HTTP-In)

#region Constrained Permissions only on server
New-PSSessionConfigurationFile -Path c:\temp\restricted.pssc -SessionType RestrictedRemoteServer -VisibleCmdlets "get-process"
notepad c:\temp\restricted.pssc
#endregion

#region double hop kerberos

# Set up variables for reuse
$ServerA = $env:COMPUTERNAME
$ServerB = Get-ADComputer -Identity ServerB
$ServerC = Get-ADComputer -Identity ServerC

# Test Kerberos double hop before changing anything -> Access is denied!
Invoke-Command -ComputerName $ServerB.Name -ScriptBlock {
    Test-Path \\$($using:ServerC.Name)\C$
}

# Notice the StartName property of the WinRM Service: NT AUTHORITY\NetworkService
# This looks like the ServerB computer account when accessing other servers over the network.
Get-WmiObject Win32_Service -Filter 'Name="winrm"' -ComputerName $ServerB.name | Select-Object SystemName, Name, StartName

# Grant resource-based Kerberos constrained delegation
Set-ADComputer -Identity $ServerC -PrincipalsAllowedToDelegateToAccount $ServerB

# Check the value of the attribute directly
$x = Get-ADComputer -Identity $ServerC -Properties msDS-AllowedToActOnBehalfOfOtherIdentity
$x.'msDS-AllowedToActOnBehalfOfOtherIdentity'.Access

# Check the value of the attribute indirectly
Get-ADComputer -Identity $ServerC -Properties PrincipalsAllowedToDelegateToAccount

# Clear the negative cache on ServerB
Invoke-Command -ComputerName $ServerB.Name -ScriptBlock {
    klist.exe purge -li 0x3e7
}

# Test Kerberos double hop again -> It Works!
Invoke-Command -ComputerName $ServerB.Name -ScriptBlock {
    Test-Path \\$($using:ServerC.Name)\C$
}

#endregion

New-PSSession -Authentication -Debug  

#CredSSP
New-PSSession -Authentication - #show options to doublehop
New-PSSessionOption -

#Is PS secure? 
Start-Process https://blogs.technet.microsoft.com/ashleymcglone/2016/06/29/whos-afraid-of-powershell-security/

#endregion

#region Module 3 Advanced Functions 1

Function Get-ServiceInfo {
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
[cmdletbinding()]
Param ($serviceNeeded='spooler',$computerName='localhost')
    Get-Service -Name $serviceNeeded -RequiredServices -ComputerName $computerName
}


function SwitchExample {
   Param([switch]$state)
   if ($state) {"on"} else {"off"}
}

function SwitchExample2 {
   Param([Bool]$state)
   if ($state) {"on"} else {"off"}
}

#dynamic parameters
Get-ChildItem -
cd cert:
Get-ChildItem -  #show the differences
Get-ChildItem -Path c:\ -

cd c:\
Get-Help Get-ChildItem -ShowWindow
cd cert:
Get-Help Get-ChildItem -ShowWindow

Update-Help
Get-Help about_functions_advanced_parameters


Function Kill-Process {
    [CmdletBinding(
                    SupportsShouldProcess=$true,
                    ConfirmImpact='Medium'
                  )
    ]

    Param([String]$Name='notepad')

    $TargetProcess = Get-Process -Name $Name
    If ($pscmdlet.ShouldProcess($name, "Stopping the process"))
        {
            $TargetProcess.Kill()
        }
} 

$ConfirmPreference = 'low'
$ConfirmPreference = 'high'
#endregion

#region Module 4 Advanced Functions 2

Function Get-ContosoObject {
Param (
  [parameter(Mandatory=$true,
             HelpMessage="Enter computer names separated by commas.")]
  [String]$ComputerName
) 
    Get-Service -ComputerName $ComputerName
}

function Get-CpuCounter {
Param (
  [ValidateSet("% Processor Time","% Privileged Time","% User Time","Mike")]
  $perfcounter
  )
Get-Counter -Counter "\Processor(_Total)\$perfcounter"
}

get-help Where-Object -ShowWindow #Number of parameter sets
get-help Set-PSReadlineOption -full

#Which verb is defined to which action
Start-Process 'https://msdn.microsoft.com/en-us/library/ms714428(v=vs.85).aspx'

#endregion

#region Module 5 Regex

#Landing page for .NET RegEx
Start-Process 'https://msdn.microsoft.com/en-us/library/hs600312(v=vs.110).aspx'
#QuickReference
Start-Process 'https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx'

#another good reference
Start-Process http://regexr.com/

#region Examples and counting multiples

"abcd defg" -match "\w+" ; $Matches
"abcd defg" -match "\W+" ; $Matches
"abcd defg" -match "\s+" ; $Matches
"abcd defg" -match "\S+" ; $Matches[1]

"abc" -match "\w{1,5}" ; $Matches
$Matches = $null

"abcd defg " -match "(?<all>(?<word1>\w{4})(?<nonword>\s)(?<word2>\w{4}))" ; $Matches
$Matches = $null

"contoso\administrator" -match "((?<domain>\w+)\\(?<user>\w+))"
$Matches.domain

"test" -match "T"
#Split/replace

'Get the keyword out of here.' -match 'keyword'

$Matches

'Get the keyword out of here.' -split 'keyword'
'Get the keyword out of here.' -replace 'keyword','otherword'

#Finding numbers in text.

$text = 'Now is the time for all good 614-555-1212men to come 123-45-6789to the aid 321-55-9999of 111-22-3333their country.'

$text -match '(?<SSN>\d{3}-\d{2}-\d{4})'

$Matches
$Matches[0]       # Overall regex match
$Matches['SSN']   # Named group

[regex]::matches($text,'(?<SSN>\d{3}-\d{2}-\d{4})') | select value #Shows multiple matches
[regex]::matches($text,'(?<SSN>\d{3}-\d{2}-\d{4})').count #Counts matches
$matches | gm

$text -match '(?<Phone>\d{3}-\d{3}-\d{4})'

$Matches
$Matches[0]       # Overall regex match
$Matches['Phone']   # Named group

#IP's
(ipconfig) -match "(?<IPv4>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"
(ipconfig) -match "(?<IPv6>\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4})"
(ipconfig) -match "(?<IPv4>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})|(?<IPv6>\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4}\:\S{0,4})"

#endregion

$text = @("www.microsoft.com", "www.Microsoft.com")
[regex]::matches($text,"[a-z]icrosoft.com")
[regex]::matches($text, "[a-z]icrosoft.com", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) 
#endregion 

#region Module 6 Error Handling

Write-Host "Hello there"
Write-Error "Error statement that I want to display"
Write-Warning "Warning error to screen"
Write-Verbose "Verbose written statement" -Verbose
Write-Debug "Debugging error to screen" -Debug
Write-Information "This is the new v5.0+ information option to screen" -InformationAction Continue

Get-Process foo
$result = Get-Process foo
$result
Get-Process >> c:\temp\process.txt

Get-Process foo 2> c:\temp\error.txt
notepad c:\temp\error.txt

Get-Process foo 2>> c:\temp\error.txt
notepad c:\temp\error.txt

$result = $null
$result = Get-Process foo,system
$result

Get-Process foo 2> c:\temp\error2.txt
notepad c:\temp\error2.txt

$result = Get-Process foo,system 2>&1
$result
$result = $null

$error[0]

#region Information stream new in PS5+

#new output stream
Write-Information "You won't see this."
Write-Information "You will see this." -InformationAction Continue

Write-Information -MessageData "This is info message" -Tags "Information" -InformationAction Continue

#Redirect stream (6 = Information)
Write-Information -MessageData "Redirect me..." -Tags "Information" 6> c:\temp\InfoStream.txt
notepad c:\temp\InfoStream.txt


#Can redirect write-host
Write-Host "Test" 6> c:\temp\WriteHostRedirect.txt
notepad C:\Temp\WriteHostRedirect.txt

$InformationPreference

Write-Host "Output is sent to host and also information steam." -InformationVariable InfoOutput
$InfoOutput

Get-Process bogus -erroraction silentlycontinue
Get-Process bogus -ea 0

#endregion

#Difference between Write-Host and Write-Output
 Function mytest  { 
 [cmdletBinding()]Param()
 Write-Host "hello";  Write-Output "good bye" 
 } 
  
$result = mytest

$result

$WindowsFolder = Get-Item C:\Windows
$?

Get-Process system,foo
$?

ping localhost -n 1
$?

ping NotValidMachineName
$?

Get-Service -WarningAction SilentlyContinue

$error.clear()
Get-Bogus
$error
$error[0]
$error | gm
$error | fl * -Force
$error[0].CategoryInfo
 
If (1 -eq 1) {
    "Line before the terminating error"
    Throw "This is my custom terminating error"
    "Line after the throw"
} 

Function SampleThrowBasedMandatoryParam {
    Param ($ComputerName = $(Throw "You must specify a value"))
    Get-CimInstance Win32_BIOS -ComputerName $ComputerName
} 

#region Trap demos

Trap [System.DivideByZeroException] {    #This is specific trap
    "Can not divide by ZERO!!!-->   " + $_.exception.message
    Continue
}
Trap  {#This is a generic catch all trap
    "A serious error occurred-->   " + $_.exception.message
    Continue
}

Write-Host -ForegroundColor Yellow "`nAttempting to Divide by Zero"
1 / $null

Write-Host -ForegroundColor Yellow "`nAttempting to find a fake file"
Get-Item c:\fakefile.txt -ErrorAction Stop

Write-Host -ForegroundColor Yellow "`nAttempting an invalid command"
1..10 | ForEach-Object {Bogus-Command} 

#endregion

$ErrorActionPreference = 0 #Can put in Try block to convert non-terminating to terminating errors


Function function3 {
    Try { NonsenseString }
    Catch {"Error trapped inside function" ; Throw}
    "Function3 was completed"
}

Try{ Function3 }
Catch { "Internal Function error re-thrown: $($_.ScriptStackTrace)" }
"Script Completed" 

type .\file.txt >> .\file.txt #This will create an endless loop until the hard drive fills up! Be careful. 

#endregion

#region Module 7 Debugging

Get-Variable #Can use in debug

#endregion

#region Module 8 DSC

#DSC future with 6.0
Start-Process https://blogs.msdn.microsoft.com/powershell/2017/09/12/dsc-future-direction-update/

Get-DscResource


Get-ScheduledTask -TaskName consistency | Start-ScheduledTask

#GPO vs. DSC
Start-Process https://blogs.technet.microsoft.com/ashleymcglone/2017/02/27/compare-group-policy-gpo-and-powershell-desired-state-configuration-dsc/

#DSC Environment Analyzer: Reporting engine module for DSC in an environment.
Start-Process https://github.com/Microsoft/DSCEA

#Lab 8.1.2 steps 7 and 10 HKLM is not correct, no 'windows' in path

#endregion

#region Module 9 Workflow


workflow test-seq {
parallel {
        sequence {
        Start-Process cmd.exe
        Get-Process -Name cmd
        }
        sequence {
        Start-Process notepad.exe
        Get-Process -Name notepad
        }
    }
}

workflow Test-WorkflowParSeq {  
 Get-Service –Name Dhcp # workflow activity
 Get-WindowsFeature –Name RSAT # Automatic InlineScript activity
 Parallel 
 {
  Sequence 
   {
   Stop-Service -Name Dhcp
   Start-Service -Name Dhcp
   }
  Sequence 
   {
   Stop-Service -Name Bits
   Start-Service -Name BITS
   }
 }
}

workflow Test-WorkflowSeqPar {

Sequence {

    Parallel {
        Stop-Service DHCP
        Stop-Service Bits
    }
    Parallel {
        Start-Service DHCP
        Start-Service Bits
    }
}
}

workflow Test-WorkflowForEach {
    ForEach -Parallel ($Item in $Items) { 
        Restart-Computer $Item
    }  
}


#Workflow demos from here:
Start-Process "https://blogs.technet.microsoft.com/heyscriptingguy/2012/12/26/powershell-workflows-the-basics/"

#Invoke-Parallel script
Start-Process https://gallery.technet.microsoft.com/scriptcenter/Run-Parallel-Parallel-377fd430/view/Discussions

#PSThreadJob, replacement option for Start-Job is Start-Thread
Start-Process https://github.com/paulhigin/psthreadjob

#region mass set home page
$path = 'HKCU:\Software\Microsoft\Internet Explorer\Main\'

$name = 'start page'

$value = 'http://blogs.technet.com/b/heyscriptingguy/'

Set-Itemproperty -Path $path -Name $name -Value $value

Function Set-HomePage {
$Computers = Get-ADComputer
$path = 'HKCU:\Software\Microsoft\Internet Explorer\Main\'
$name = 'start page'
$value = 'http://blogs.technet.com/b/heyscriptingguy/'
    Workflow Reset-AllComputerHomePages {
        Parallel {            
                    Invoke-Command -ComputerName $Computers -ScriptBlock {Set-Itemproperty -Path $path -Name $name -Value $value}
                 }
     }
}

#endregion

#endregion

#region A2 Jobs

Get-ChildItem c:\ -Recurse
Start-Job -Name GetAllFiles -ScriptBlock {Get-ChildItem c:\ -Recurse}

Get-Job | Stop-Job
Receive-Job -Name GetAllFiles -Keep

Get-Job -Name GetAllFiles | Remove-Job


#endregion

#endregion

#region Make PowerShell work better for you

#Ashley McGlone
Start-Process https://blogs.technet.microsoft.com/ashleymcglone/2017/07/12/slow-code-top-5-ways-to-make-your-powershell-scripts-run-faster/

#Dan Sheehan
Start-Process https://blogs.technet.microsoft.com/heyscriptingguy/2014/05/18/weekend-scripter-powershell-speed-improvement-techniques/


#endregion

#region PS Script Analyzer
Invoke-ScriptAnalyzer -Path 

Invoke-ScriptAnalyzer -Path 'C:\Users\mconeill\OneDrive - Microsoft\_Modules and Scripts\O365_Logon Module\v1.1\O365_Logon\O365_Logon.psm1' -ExcludeRule PSAvoidUsingWriteHost

#Regular expression search for non-ASCII characters: [^\u0000-\u007F]

#$null should be on the left side of equality comparisons
Start-Process https://github.com/PowerShell/PSScriptAnalyzer/issues/200
Start-Process https://connect.microsoft.com/PowerShell/feedback/details/1299466/pspossibleincorrectcomparisonwithnull-null-should-be-on-the-left-side-of-equality-comparisons
#PowerShell's comparison operators return an array of items if the lhs is an array rather than a [bool]. e.g
#endregion

#region New to PS 5

#region zip files
Compress-Archive

Compress-Archive -Path C:\Temp\* -CompressionLevel Fastest -DestinationPath C:\Temp\zipped

Expand-Archive
Expand-Archive -LiteralPath C:\Temp\Zipped.Zip -DestinationPath C:\Temp\UnZipped_Folder
#endregion

#region New-TemporaryFile
$temp = New-TemporaryFile
$temp

Add-Content -Path $temp -Value "Hello, World"
Add-Content -Path $temp -Value "Hello, World"
Add-Content -Path $temp -Value "Hello, World"
Add-Content -Path $temp -Value "Hello, World" -NoNewline 
Add-Content -Path $temp -Value "Hello, World" -NoNewline
Add-Content -Path $temp -Value "Hello, World" -NoNewline

Get-Content -Path $temp
$temp
Get-Command -ParameterName NoNewLine

Remove-Item $temp
#endregion

#region New-Guid
New-Guid
New-Guid | gm
$g = New-Guid
$g
$g | gm
$g.Guid

#endregion

#region Recycle bin
Clear-RecycleBin -DriveLetter c

#endregion

#endregion

#region Desired State Configuration Workshop

#region Knowledge base
Start-Process https://aka.ms/DSCEA #DSC reporting dashboard
#endregion

#region Module 1 Introduction

#DSC vs. GPO
Start-Process https://blogs.technet.microsoft.com/ashleymcglone/2017/02/27/compare-group-policy-gpo-and-powershell-desired-state-configuration-dsc/

#What is Dev/Ops
Start-Process aka.ms/devopsforn00bs

#region DSC demo
#Requires -RunAsAdministrator 

Set-Location C:\temp

# The module that makes DSC possible
Get-Command -Module PSDesiredStateConfiguration

# Engine status
Get-DscLocalConfigurationManager

# No configuration applied
Get-DscConfiguration

# CTRL-J and select DSC Configuration (simple)
# Use CTRL-SPACE to invoke Intellisense on the resource keywords to find out their syntax

Configuration MyFirstConfig
{
    Node localhost
    {
        Registry RegImageID {
            Key = 'HKLM:\Software\Contoso'
            ValueName = 'ImageID'
            ValueData = '42'
            ValueType = 'DWORD'
            Ensure = 'Present'
        }

        Registry RegAssetTag {
            Key = 'HKLM:\Software\Contoso'
            ValueName = 'AssetTag'
            ValueData = 'A113'
            ValueType = 'String'
            Ensure = 'Present'
        }

        Registry RegDecom {
            Key = 'HKLM:\Software\Contoso'
            ValueName = 'Decom'
            ValueType = 'String'
            Ensure = 'Absent'
        }

        Service Bits {
            Name = 'Bits'
            State = 'Running'
        }

    }
}

# Generate the MOF
MyFirstConfig

# View the MOF
Get-ChildItem .\MyFirstConfig
notepad .\MyFirstConfig\localhost.mof

# Check state manually
Get-ItemProperty HKLM:\Software\Contoso\
Get-Service BITS

# Check state with cmdlet
Test-DscConfiguration

# Sets it the first time
Start-DscConfiguration -Wait -Verbose -Path .\MyFirstConfig

# Check state manually
Get-Item HKLM:\Software\Contoso\
Get-Service BITS

# View the config of the system
Get-DscConfiguration

# Check state with cmdlet
Test-DscConfiguration

# Change the state
Set-ItemProperty HKLM:\Software\Contoso\ -Name ImageID -Value 12
New-ItemProperty HKLM:\Software\Contoso\ -Name Decom -Value True
Stop-Service Bits

# Check state manually
Get-Item HKLM:\Software\Contoso\
Get-Service BITS

# View the config of the system
Get-DscConfiguration

# Do I have the registry key? Is the value correct?
Test-DscConfiguration

# Reset the state
Start-DscConfiguration -Wait -Verbose -Path .\MyFirstConfig

# Check state manually
Get-Item HKLM:\Software\Contoso\
Get-Service BITS

# View the config of the system
Get-DscConfiguration

# Check state with cmdlet
Test-DscConfiguration




# Reset demo
Remove-Item C:\Windows\System32\Configuration\Current.mof, C:\Windows\System32\Configuration\backup.mof, C:\Windows\System32\Configuration\Previous.mof
Remove-Item HKLM:\Software\Contoso
Remove-Item .\MyFirstConfig -Recurse -Force -Confirm:$false
Stop-Service Bits

#endregion

#My first configuration
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\00-MyFirstConfig.ps1'

#endregion

#region Module 2 Push



#Labs a-d localhost
#Bogus MOF
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\02a-MOFs.ps1'

#LCM
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\02b-LCM.ps1'

#configurations
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\02c-Configurations.ps1'
Get-DscResource windowsoptionalfeature -Syntax
#Push
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\02d-Push.ps1'

#labs e-f in lab

#endregion

#region Module 3 Pull

#Configuration
ise 'C:\Users\mconeill\OneDrive - Microsoft\Workshops\PowerShell Desired State Configuration\Demos\DemoScriptsV1.2\03b-Config2.ps1'

#endregion

#region Module 4 Security

#endregion

#region Module 5 Resources

Get-DscResource

Get-DscResource -Name user -Syntax

#region xPendingReboot
Get-DscResource -Name xPendingReboot #confirm if module is installed or not

Install-Module xPendingReboot #Run as Admin

Configuration CheckForPendingReboot
{       
  Import-DscResource -module xPendingReboot
    Node ‘localhost’
    { 
        xPendingReboot Reboot1
        {
            Name = ‘BeforeSoftwareInstall’
        }
        LocalConfigurationManager
        {
            RebootNodeIfNeeded = $true
        }
    } 
}
CheckForPendingReboot -OutputPath c:\temp\
Start-DscConfiguration -Path C:\Temp\ -Wait -Verbose

Test-DscConfiguration
Get-DscLocalConfigurationManager
Get-DscResource xPendingReboot -Syntax

#endregion

#Updated naming's
Start-Process https://blogs.msdn.microsoft.com/powershell/2017/12/08/dsc-resource-naming-and-support-guidelines/ 

Find-DscResource

#endregion

#region Module 6 Custom Resources
dir variable: | out-file c:\temp\log.txt
psedit C:\temp\log.txt
#endregion

#region Module 7 Troubleshooting
dir c:\windows\System32\Configuration #look at LCM
dir C:\windows\System32\Configuration\ConfigurationStatus #look at LCM logs
#endregion

#region Appendix 1 Advanced Configuration

#endregion

#region Appendix 2 LCM Scenarios

#endregion

#region Appendix 3 Reporting

#endregion

#region Appendix 4 Azure

#endregion

#endregion

#region to work on 
$ComputersWanted = '2012R2-DC' ,'pc2','pc3'

Foreach ($ComputerWanted in $computersWanted) {
write-host "starting process on $computerwanted "
Restart-Computer -ComputerName $ComputerName
Start-Sleep -Seconds 30

while (-not (Get-Service BITS -ComputerName $ComputerName  | Where-Object {$_.Status -eq 'running'} ))
{
    "Waiting on restart: $ComputerName"
    Start-Sleep -Seconds 30
} 

Write-Host "Computer is active: $computerwanted"

}




#endregion