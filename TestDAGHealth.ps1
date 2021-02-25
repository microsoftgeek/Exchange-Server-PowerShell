#This Sample Code is provided for the purpose of illustration only
#and is not intended to be used in a production environment.  THIS
#SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT
#WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
#LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
#FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free
#right to use and modify the Sample Code and to reproduce and distribute
#the object code form of the Sample Code, provided that You agree:
#(i) to not use Our name, logo, or trademarks to market Your software
#product in which the Sample Code is embedded; (ii) to include a valid
#copyright notice on Your software product in which the Sample Code is
#embedded; and (iii) to indemnify, hold harmless, and defend Us and
#Our suppliers from and against any claims or lawsuits, including
#attorneys' fees, that arise or result from the use or distribution
#of the Sample Code.
#
#
# -----------------------------------------------------------------------
# This script shows current status of DAG databases.
# -----------------------------------------------------------------------
#
#

(Get-DatabaseAvailabilityGroup -Identity (Get-MailboxServer -Identity $env:computername).DatabaseAvailabilityGroup).Servers | Test-MapiConnectivity | Sort Database | Format-Table -AutoSize
Get-MailboxDatabase | Sort Name | Get-MailboxDatabaseCopyStatus | Format-Table -AutoSize
function CopyCount 
{
$DatabaseList = Get-MailboxDatabase | Sort Name
$DatabaseList | % {
$Results = $_ | Get-MailboxDatabaseCopyStatus
$Good = $Results | where { ($_.Status -eq "Mounted") -or ($_.Status -eq "Healthy") }
$_ | add-member NoteProperty "CopiesTotal" $Results.Count
$_ | add-member NoteProperty "CopiesFailed" ($Results.Count-$Good.Count)
}
$DatabaseList | sort copiesfailed -Descending | ft name,copiesTotal,copiesFailed -AutoSize 
}
CopyCount
