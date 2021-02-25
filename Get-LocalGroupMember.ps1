
#region One way to get local admins 

#Written 4/22/2011 Author TFosler@Microsoft Corp.
#modified 10/11/2016 - belder@microsoft.com
# This script will allow you to remotely list members of the local Administrators group

#set output path
$path= 'c:\bobtestfiles\dump'
Remove-Item -path $path\serversandlocaladmins.txt -Force

"{0}`,`{1}" -f "Server","Acct" | Out-File $path\serversandlocaladmins.txt

#Array of Computer Names that Need Local Admin Password Changed
$strComputers = Get-Content $path\servers.txt
# @("dept-wrks1.my.domain.net","dept-wrks2.my.domain.net","dept-wrks3..my.domain.net")
# test connection to server before attempting changes.
[int]$port=445
foreach($strComputer in $strComputers)
{
  
  $ErrorActionPreference = “SilentlyContinue”

  $socket = new-object Net.Sockets.TcpClient

  $socket.Connect($strComputer, $port)

if ($socket.Connected) 
{ $socket.Close()
$computer = [ADSI]("WinNT://" + $strComputer + ",computer")
    $group = $Computer.psbase.children.find("Administrators") # Change the Administrators group to another local group if you want to list members of that group
    $group.Name

    

    function ListAdministrators # this will also list the members that currently exist, but also shows you that the account has been added in the results.
    {$members = $group.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name",'GetProperty',$null,$_,$null)}
    $members}
    ListAdministrators
    foreach ($box in ListAdministrators){
    "{0}`,`{1}" -f $strComputer, $box | out-file $path\serversandlocaladmins.txt -Append
    }

    #Write-Host $strComputer
    

}
Else
{"{0}`,`{1}" -f "ping failed",$strComputer| Out-File $path\connectfailure.txt -append}
  $socket = $null
}

import-csv $path\serversandlocaladmins.txt -delimiter "`," | export-csv -path $path\csvfile.csv -NoTypeInformation 
#endregion

#region a second way to get local admins

<#
.Synopsis
   Show members of a local group in the Targeted computername
.DESCRIPTION
Show members of a local group in the Targeted computername
.EXAMPLE
   Get-LocalGroupMember -name TestGroup -computername remotepc1
#>
function Get-LocalGroupMember
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string[]]$Name,
        
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [string[]]$Computername="$ENV:Computername"

        
    )

    Begin
    {
    }
    Process
    {
    # Code for decoding group membership provided
    # Courtesy of Francois-Xaver Cat 
    # Windows PowerShell MVP
    # Thanks Dude!
    $group = [ADSI]"WinNT://$($computername[0])/$($Name[0]),group" 
    $member=@($group.psbase.invoke("Members"))
    $member | ForEach-Object {([ADSI]$_).InvokeGet("Name")}
        
    }
    End
    {
    }
}

#endregion

#region PS 5.1+ way to get local admins

Get-LocalGroupMember

#endregion