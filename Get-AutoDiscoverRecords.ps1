<#
.NOTES
	Name: Get-AutoDiscoverRecords.ps1
	Author: Daniel Sheehan
	Requires: PowerShell v4 (the Exchange Management Shell is not required).
	Version History:
	1.0 - 6/22/2016 - Initial Release.
	1.1 - 9/19/2016 - Switched to longer date/time string and added logic to
	interpret it according to the local region settings of the computer running
	the script.
	1.2 - 9/22/2016 - Added check to see if the sites listed in the SCP record
	are current AD sites when the ADSite "*" is specified.
	###########################################################################
	The sample scripts are not supported under any Microsoft standard support
	program or service. The sample scripts are provided AS IS without warranty
	of any kind. Microsoft further disclaims all implied warranties including,
	without limitation, any implied warranties of merchantability or of fitness
	for a particular purpose. The entire risk arising out of the use or
	performance of the sample scripts and documentation remains with you. In no
	event shall Microsoft, its authors, or anyone else involved in the
	creation, production, or delivery of the scripts be liable for any damages
	whatsoever (including, without limitation, damages for loss of business
	profits, business interruption, loss of business information, or other
	pecuniary loss) arising out of the use of or inability to use the sample
	scripts or documentation, even if Microsoft has been advised of the
	possibility of such damages.
	###########################################################################
.SYNOPSIS
	Provides a list of all found AutoDiscover SCP Record information found in
	AD.
.DESCRIPTION
	This script gathers all Exchange AudtoDiscover SCP records defined in AD,
	and provides the AutoDiscover information for the specified site or all
	sites depending on the parameters provided. The information returned
	includes the server name the SCP record belongs to, the site name it
	covers, whether the server covers multiple sites, the date it was created,
	and the URL it points to. If a server's SCP record covers multiple sites,
	each site will be listed separately for sorting purposes, but the entires
	will reflect a MultiSite value of "True" versus "False". If the ADSite "*"
	was specified, then each site listed in a SCP record is checked against
	current sites in AD, with the ValidSite value of "True" or "False" is
	recorded. The output is either a CSV file if specified, otherwise it is
	displayed as a table in the PowerShell window.
	This script is based on the code found here:
	https://vanhybrid.com/2012/11/21/retrieving-exchange-autodiscover-scp-information-from-ad-via-powershell/
	DateTime conversion piece is based on the code found here:
	http://www.powershellmagazine.com/2013/07/08/pstip-converting-a-string-to-a-system-datetime-object/
.PARAMETER ADSite
	This optional switch tells the script which AD site to query the Exchange
	servers to generate the report.
.PARAMETER OutCSVFile
	This optional parameter specifies the path and file name of the CSV to
	export the gathered data to. If this parameter is omitted then the
	information is printed to the screen.
.EXAMPLE
	[PS] C:\>.\Check-ExchangeVersion.ps1 <no parameters>
	The AutoDiscover SCP record information for the computer's current local
	site is output to the screen in a table format.
.EXAMPLE
	[PS] C:\>.\Check-ExchangeVersion.ps1 -ADSite MD-Rockville
	The AutoDiscover SCP record information for the MD-Rockville site is
	output to the screen in a table format.
.EXAMPLE
	[PS] C:\>.\Check-ExchangeVersion.ps1 -ADSite *
	-OutCSVFile .\SCPInfoAllSites.CSV
	All AutoDiscover SCP record information is exported to the
	SCPInfoAllSites.CSV file in the current directory, including if the listed
	sites are valid AD sites or not.
.LINK
	https://gallery.technet.microsoft.com/Get-AutoDiscover-Records-86db854c
#>

Param (
	[Parameter(Mandatory = $False)]
	# Default to the current local site the script is run in if none is specified.
	[String]$ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name,
	[Parameter(Mandatory = $False)]
	[String]$OutCSVFile
)

# Extract the default Date/Time formatting from the local computer's "Culture" settings, and then create the format to use when parsing the date/time information pull from AD.
$CultureDateTimeFormat = (Get-Culture).DateTimeFormat
$DateFormat = $CultureDateTimeFormat.ShortDatePattern
$TimeFormat = $CultureDateTimeFormat.LongTimePattern
$DateTimeFormat = "$DateFormat $TimeFormat"

# Add a blank line to make the script output stand out.
Write-Host ""

# Check to see if the ADSite parameter was set to "*", and if it was then pull all existing AD site link names into an array.
If ($ADSite -eq "*") {
	Write-Host "Gathering current Site names from AD for SCP record comparison. This could take some time in large environments."
	# Create an AD query using native objects calls so the ActiveDirectory Module doesn't have to be used.
	$SiteSearch = New-Object System.DirectoryServices.DirectorySearcher
	# Set the filter to all AD Site objects.
	$SiteSearch.Filter = '(objectClass=site)'
	# Set the root to the Configuration Container under the forest root name space.
	$SiteSearch.SearchRoot = "LDAP://" + ([ADSI]"LDAP://RootDSE").configurationNamingContext
	$FoundSites = $SiteSearch.FindAll()
	# Store each AD Site Name in an array named SiteArray.
	$SiteArray = @()
	ForEach ($FoundSite in $FoundSites) {
		$SiteArray += $FoundSite.Properties.name
	}
	Write-Host "Gathering all AutoDiscover SCP record information."
} Else {
	Write-Host "Gathering AutoDiscover SCP record information for the AD Site: " -NoNewline
		Write-Host "$ADSite" -ForegroundColor Green
}

# Create the empty array to store the SCP entries.
$SCPEntries = @()

# Create an AD query using native objects calls so the ActiveDirectory Module doesn't have to be used.
$SCPSearch = New-Object System.DirectoryServices.DirectorySearcher
# Set the filter to all SCP objects with one of two keywords.
$SCPSearch.Filter = '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))'
# Set the root to the Configuration Container under the forest root name space.
$SCPSearch.SearchRoot = "LDAP://" + ([ADSI]"LDAP://RootDSE").configurationNamingContext

# Loop through each SCP entry found in the query.
ForEach ($SCPEntry in $SCPSearch.FindAll()) {
	# Pull up the record's properties.
	$SCPRecord = [ADSI]$SCPEntry.Path
	# Each SCP record should have 2 or more keywords, one of the keywords strings above and one "Site=<SiteName>".
	# Check to see if there was at least 1 keyword.
	If ($SCPRecord.Keywords.Count -gt 1) {
		# There was so record the True/False check of if there were more than 2 keywords, which means there are multiple AD Sites covered by the server.
		[String]$MultiSite = ($SCPRecord.Keywords.Count -gt 2)

		# Loop through each of the keywords in the record.
		ForEach ($SCPSite in $SCPRecord.Keywords) {
			# For each Site= entry, create a new object/entry in the array for the server.
			# This means if a server covers multiple sites, it will be listed multiple times in the array with the MultiSite attribute of "Yes".
			If ($SCPSite -like "Site=*") {
				# Extract the site name by dropping the first 5 characters as they are always "Site=" at this point.
				$SCPSiteName = $SCPSite.SubString(5)
				$SCPData = New-Object PSObject -Property @{
					Server = $SCPRecord.cn.ToString()
					Site = $SCPSiteName
					MultiSite = $MultiSite
					DateCreated = [DateTime]::ParseExact($SCPRecord.WhenCreated.ToString(),$DateTimeFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)
					AutoDiscoverInternalURI = $SCPRecord.ServiceBindingInformation.ToString()
				}
				# If all AD sites are queried, then record the True/False check if the site listed in the SCP record currently exists as an AD site name.
				If ($ADSite -eq "*") {
					$SCPData | Add-Member -Type NoteProperty -Name ValidSite -Value $([String]($SiteArray -contains $SCPSiteName))
				}
				$SCPEntries += $SCPData
			}
		}
	} Else {
		# Otherwise if there was only 1 keyword, then that means the site coverage attribute on the CAS is blank.
		Write-Warning "The Server $($SCPRecord.cn.ToString()) does not list any AD sites for AutoDiscover!"
	}
}

# Check to see if the ADSite value of "*" was used, which means all sites should be returned.
If ($ADSite -eq "*") {
	# It was so include all the site entries, and sort them first by Site name and then by date created.
	$AutoDiscoverRecords = $SCPEntries | Sort -Property DateCreated,Site | Select Server,Site,MultiSite,DateCreated,ValidSite,AutoDiscoverInternalURI
} Else {
	# It wasn't so just include the entires with the matching ADSite name, and sort them by date created.
	$AutoDiscoverRecords = $SCPEntries | Where {$_.Site -like $ADSite} | Sort -Property DateCreated | Select Server,Site,MultiSite,DateCreated,AutoDiscoverInternalURI
}

# Check to make sure there was at least 1 Auto Discover SCP records were returned with the supplied ADSite filter.
If ($AutoDiscoverRecords.Count -eq 0) {
	# There were 0 rows so report that and do nothing else.
	Write-Host -ForegroundColor Red "`nNo site data was found with the AD Site `"$ADSite`", so there is nothing to report/export."
# There was 1 or more records found, so check to see the OutCSVFile parameter was used.
} ElseIf ($OutCSVFile) {
	# It was so export the data to the specified CSV.
	$AutoDiscoverRecords | Export-CSV -NoTypeInformation $OutCSVFile
	Write-Host -ForegroundColor Green "The collected SCP data was exported to the `"$OutCSVFile`" CSV file."
} Else {
	# It wasn't so display the information to screen.
	Write-Host -ForegroundColor Green "The collected SCP data is as follows:"
	$AutoDiscoverRecords | Format-Table -AutoSize
}