<#
    DISCLAIMER: This application is a sample application. The sample is provided "as is" without 
    warranty of any kind. Microsoft further disclaims all implied warranties including without 
    limitation any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the samples remains with you. 
    In no event shall Microsoft or its suppliers be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss arising out of the use of or inability 
    to use the samples, even if Microsoft has been advised of the possibility of such damages. 
    Because some states do not allow the exclusion or limitation of liability for consequential or 
    incidental damages, the above limitation may not apply to you.

    ************************************
    Created by: Hector Ventura
    E-mail: hventura@aceitconsultinginc.com
    ************************************
    Move all emails from one mailbox folder to another in the main mailbox. Also you can use this script to archive
    emails moving all items from one folder to another folder located inside  then "In-Place Archive". See examples below

    ************************************
    Prerequisites
    ************************************
    1 - The script requires EWS Managed API 2.2, which can be downloaded here: 
    https://www.microsoft.com/en-gb/download/details.aspx?id=42951

    2 - Administrator needs to have Impersonation Role assigned.
    To assign application impersonation management role in Office 365, you need to follow the stepwise instruction mentioned below:
    1.Using Exchange Admin Center or an admin account, log in to your Microsoft Office 365 account 
    2.Now, on Office 365 access the Exchange tab and go to Permissions in the left pane under Dashboard 
    3.After that click on admin roles and then select Discovery management by double-clicking it in the right pane 
    4.In Discovery Management Window, click on + button to set application impersonation 
    5.Now, from the drop down list select "ApplicationImpersonation" and click on add button then, click on OK button 
    6.To verify, check ApplicationImpersonation has been added under the roles or not 
    7.Now, go to Members section and click on + option, a new Window get appears 
    8.Choose the user name and click on add button --> click OK 
    9.To verify, check the Member section for the user name, if user name is in the list click on Save button, otherwise go to step 7 

    3 - TargetFolder has to be created previous to run the Script

    4 - Source and Target folders have to be unique names. No repeated folders in subfolders. 

    ************************************

    Use of the script:

    Move all emails from inbox to archive folder in the primary mailbox
    PS>.\MoveEmails.ps1 -MailboxName test@o365genius.com -SourceFolder Inbox -SourceType primary -TargetFolder BackupFolder -TargetType primary -username "admin@domain.com" -password "p@ssW0rd"
 

    Move all emails from inbox/main mailbox to Inbox/archive folder in the archive online. It will ask for username and password.
    PS>.\MoveEmails.ps1 -MailboxName test@o365genius.com -SourceFolder Inbox -SourceType primary -TargetFolder Inbox -TargetType archive 


    Move all emails from archive folder to Test folder in the archive online. It will ask for admin username and password.
    PS>.\MoveEmails.ps1 -MailboxName test@o365genius.com -SourceFolder Archive -SourceType archive -TargetFolder test -TargetType archive

#>

param (
  [Parameter(Position=0,Mandatory=$True,HelpMessage='Specifies the mailbox to be accessed')]
  [ValidateNotNullOrEmpty()]
  [string]$MailboxName,
	
  [Parameter(Position=1,Mandatory=$True,HelpMessage='Source folder (from which to move messages)')]
  [string]$SourceFolder,
	
  [Parameter(Position=2,Mandatory=$True,HelpMessage='is the Source folder primary or archive?')]
  [string]$SourceType,

  [Parameter(Position=3,Mandatory=$True,HelpMessage='Target folder (messages will be moved here)')]
  [string]$TargetFolder,

  [Parameter(Position=4,Mandatory=$True,HelpMessage='is the target folder primary or archive')]
  [string]$TargetType,

  [Parameter(Position=5)]
  [string]$username,

  [Parameter(Position=6)]
  [string]$password
	
);

[string]$warning = 'Yellow'                      # Color for warning messages
[string]$myerror = 'Red'                           # Color for error messages
[string]$LogFile = '.\Log.txt'             # Path of the Log File


$icount = 0;
$i=0;
$FolderId = @();
$FolderName = "$SourceFolder","$TargetFolder"
$Location = "$SourceType","$TargetType"

#if username or password are empty, ask for both
if ([string]::IsNullOrEmpty($username) -or [string]::IsNullOrEmpty($password)) {
  $p = get-credential -Message 'Admin Username and password with impersonation permission'
  if ($p) {
    $username = $p.UserName
    $BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($p.Password)
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  }
  else { 
    Write-Error -Message 'Admin credential needed'
    return $False
  }
}


# Make sure the Import-Module command matches the Microsoft.Exchange.WebServices.dll location of EWS Managed API, chosen during the installation
Import-Module -Name 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

#Creating the Exchange service object
$service = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2013_SP1

#Provide the credentials of the O365 account that has impersonation rights on the mailbox $MailboxName
$service.Credentials = new-object -TypeName Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $username, $password

#Exchange Online URL
$service.Url= new-object -TypeName Uri -ArgumentList ('https://outlook.office365.com/EWS/Exchange.asmx')

#User to impersonate
$service.ImpersonatedUserId = new-object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxName)
$service.HttpHeaders.Add('X-AnchorMailbox', $MailboxName)

While ($i -lt 2) {
  if ($Location[$i] -eq 'primary') {
    $FolderView = new-object -TypeName Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList (100)
    $FolderView.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep
    $SearchFilter = new-object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList ([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName[$i])
    $FindFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SearchFilter,$FolderView)

    if($FindFolderResults.Id) {
      Write-host 'The folder' $FolderName[$i] 'was successfully found in the primary mailbox' -ForegroundColor $warning
      $FolderId += $FindFolderResults.Id
      $i++;
      continue;
    }
    else {  
      Write-host 'The folder' $FolderName[$i] 'was not found in the primary mailbox' -ForegroundColor $warning; $i++; continue;
    }
  }
  else {
    $j=0;
    $Mbx = (Get-Mailbox $MailboxName)
    if ($Mbx.ArchiveStatus -eq 'Active') {  
      $guid=($Mbx.ArchiveGuid).ToString();
      (Get-MailboxFolderStatistics $guid).Name | ForEach-Object { $j++;
        if ($FolderName[$i] -match $_) {
          Write-host 'The folder' $FolderName[$i] 'was successfully found in the archive mailbox' -ForegroundColor $warning
          $AFolderView = new-object -TypeName Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList (100)
          $AFolderView.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep
          $ASearchFilter = new-object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList ([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName[$i])
          $AFindFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$ASearchFilter,$AFolderView)
          $FolderId += $AFindFolderResults.Id
          $i++;
          continue;
        }
        else { 
          if ($j -eq (Get-MailboxFolderStatistics $guid).Count) { 
            Write-Host 'All the archive folders have been checked and' $FolderName[$i] 'was not found' -ForegroundColor $warning; $i++; continue;
          } 
        }
      }
    }
    else { 
    Write-host 'The archive mailbox is not enabled.' -ForegroundColor $warning; $i++; continue; }
  }
}

if($FolderId.Count -eq 2)  {
  $ItemView = new-object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList (1000)
  $icount = 1
  do {
    $FindItemResults = $service.FindItems($FolderId[0],$ItemView)
    write-host $FindItemResults.TotalCount 'items have been found in the Source folder and will be moved to the Target folder.'
    foreach ($Item in $FindItemResults.Items) {
      $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$Item.Id)
      $Message.Move($FolderId[1]) > $null

      if (($icount % 100) -gt 0) { write-host '.' -NoNewline }
      else { write-host ('{0}' -f $icount) }
            
      $icount += 1
    }
    $ItemView.offset += $FindItemResults.Items.Count
  } while($FindItemResults.MoreAvailable -eq $true)

  if (($icount % 100) -gt 0) { write-host ($icount-1) }
}
else 
{
  if ($FolderId.Count -gt 2) {
    Write-Host 'Either the Source or Target folder is repeated. Rename to unique names' -ForegroundColor $Myerror 
  }
  else {  
    Write-host 'Check the source and the target folders. One of them is probably invalid.' -ForegroundColor $myerror 
  }
}

#Catch the errors
trap [Exception]
{
  Write-host ('Error: ' + $_.Exception.Message) -foregroundcolor $myerror;
  Add-Content -Path $LogFile -Value ('Error: ' + $_.Exception.Message);
  continue;
}