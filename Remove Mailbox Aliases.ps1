$Mailboxes = Get-Mailbox -ResultSize Unlimited -OrganizationalUnit Accounts

 

ForEach($Mailbox in $Mailboxes)

{

  $Mailbox.EmailAddresses |

  ?{$_.AddressString -like '*@email.lifetouch.com'} |

  %{Set-Mailbox $Mailbox -EmailAddresses @{remove=$_}

  }

}