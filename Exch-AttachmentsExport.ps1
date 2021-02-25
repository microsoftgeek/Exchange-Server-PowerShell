Write-Output "$HR Copy Email Attachments to Network Share

##########################################################################################
#
#                  *CDI - Copy Email Attachments to Network Share* 
#                                                                                
# Created by Cesar Duran (Jedi Master)                                                                                        
# Version:1.0                                                                                                                                       
#                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
#                                                                                                                                                                                                          
###########################################################################################

$HR"

# CDI - Copy Email Attachments to Network Share

# Line delimiter
$HR = "`n{0}`n" -f ('='*20)


########################################
Write-Output "$HR Connecting to EXO via the InterWebs $HR"
##-----------------------------------------------------##
## Connect to Exchange Online                          ##
##-----------------------------------------------------##

# Store your credentials - Enter you username and the app password
$Credentials = Get-Credential
$Credentials.Password | ConvertFrom-SecureString | Set-Content C:\test\password.txt
$Username = $Credentials.Username
$Password = Get-Content “C:\test\password.txt” | ConvertTo-SecureString
$Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password


# Connect to Msol
Connect-MsolService -Credential $Credentials

# Connect to AzureAd
Connect-AzureAD -Credential $Credentials

# Connect to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking



########################################
Write-Output "$HR Release the Kraken! $HR"
##-----------------------------------------------------##
## Email Attachment Script Below                       ##
##-----------------------------------------------------##

# Name of the mailbox to pull attachments from
$MailboxName = 'cesar-test@cdirad.com'


# Location to move attachments
$downloadDirectory = '\\mn-sl-dfs-2\departments\IT\SystemAdmins\email-attachments\'


# Path to the Web Services dll
$dllpath = "D:\Program Files\Microsoft\Exchange Server\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[VOID][Reflection.Assembly]::LoadFile($dllpath)


# Create the new web services object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)


# Create the LDAP security string in order to log into the mailbox
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind


# Auto discover the URL used to pull the attachments
$service.AutodiscoverUrl($aceuser.mail.ToString())


# Get the folder id of the Inbox
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)


# Find mail in the Inbox with attachments
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfha)


# Grab all the mail that meets the prerequisites
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)


# Loop through the emails
foreach ($miMailItems in $frFolderResult.Items){

	# Load the message
	$miMailItems.Load()

	# Loop through the attachments
	foreach($attach in $miMailItems.Attachments){

		# Load the attachment
		$attach.Load()

		# Save the attachment to the predefined location
		$fiFile = new-object System.IO.FileStream(($downloadDirectory + “\” + (Get-Date).Millisecond + "_" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
		$fiFile.Write($attach.Content, 0, $attach.Content.Length)
		$fiFile.Close()
	}

	# Mark the email as read
	$miMailItems.isread = $true
	$miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)

	# Delete the message (optional)
	[VOID]$miMailItems.Move("DeletedItems")
}