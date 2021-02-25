﻿<#
.NOTES
	Name: HealthChecker.ps1
	Author: Marc Nivens
	Requires: Exchange 2013 Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
	Version History:
	1.22 - 2/9/2015
	3/30/2015 - Initial Public Release.
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Checks the target Exchange server for various configuration recommendations from the Exchange product group.
.DESCRIPTION
	This script checks the Exchange server for various configuration recommendations outlined in the 
	"Exchange 2013 Performance Recommendations" section on TechNet, found here:

	https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx

	Informational items are reported in Grey.  Settings found to match the recommendations are
	reported in Green.  Warnings are reported in yellow.  Settings that can cause performance
	problems are reported in red.  Please note that these recommendations only apply to Exchange
	2013.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the 		
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local 		
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on 		
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER MailboxReport
	This optional parameter gives a report of the number of active and passive databases and
	mailboxes on the server.
.PARAMETER LoadBalancingReport
    This optional parameter will check the connection count of the Default Web Site for every server
    running Exchange 2013 with the Client Access role.  It then breaks down servers by percentage to 
    give you an idea of how well the load is being balanced.
.PARAMETER CasServerList
    Used with -LoadBalancingReport.  A comma separated list of CAS servers to operate against.  Without 
    this switch the report will use all 2013 Client Access servers in the organization.
.PARAMETER Verbose	
	This optional parameter enables verbose logging.
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME -MailboxReport -Verbose
	Run against a single remote Exchange server with verbose logging and mailbox report enabled.
.EXAMPLE
    Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}
    Run against all Exchange 2013 servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport
    Run a load balancing report comparing all Exchange 2013 CAS servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport -CasServerList CAS01,CAS02,CAS03
    Run a load balancing report comparing servers named CAS01, CAS02, and CAS03.
.LINK
    https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx
    https://technet.microsoft.com/en-us/library/36184b2f-4cd9-48f8-b100-867fe4c6b579(v=exchg.150)#BKMK_Prereq
#>

# Use the CmdletBinding function so the script accepts and understands -Verbose and -Debug and sets the default parameter set to
# "Gather". The Write-Verbose and Write-Debug statements in this script will activate only if their respective switches are used.
[CmdletBinding(DefaultParameterSetName = "Gather")]

#Parameters
param
(
    #user local computer name if no name is specified
    $Server = ($env:COMPUTERNAME),
    #User specified path should not end in a \
    [ValidateScript({-not $_.ToString().EndsWith('\')})]$OutputFilePath = ".",
    [switch]$MailboxReport,
    [switch]$LoadBalancingReport,
    $CasServerList = $null,
	[switch]$ServerReport,
	$ServerList
)

# Check to see if the -Verbose parameter was used.
If ($PSBoundParameters["Verbose"]) {
	#Write verbose output in Cyan since we already use yellow for warnings
	$VerboseForeground = $Host.PrivateData.VerboseForegroundColor
	$Host.PrivateData.VerboseForegroundColor = "Cyan"
}

#Enums and custom data types
Add-Type -TypeDefinition @"
    namespace HealthChecker
    {
        public enum ServerRole
        {
            MultiRole,
            Mailbox,
            ClientAccess,
            Edge,
            None
        }
        public enum ServerType
        {
            VMWare,
            HyperV,
            Physical,
            Unknown
        }
        public enum OSVersion
        {
            Windows2008,
            Windows2008R2,
            Windows2012,
            Windows2012R2,
            Unknown
        }
    }
"@

$script:ServerResultList = New-Object System.Collections.ArrayList

#Versioning
$ScriptName = "Exchange 2013 Health Checker"
$ScriptVersion = "1.22"
$OutputFileName = "HealthCheck" + "-" + $Server + "-" + (get-date).tostring("MMddyyyyHHmmss") + ".log"
$OutputFullPath = $OutputFilePath + "\" + $OutputFileName
$VirtualizationWarning = @"
Virtual Machine detected.  Certain settings about the host hardware cannot be detected from the virtual machine.  Verify on the VM Host that: 

    - There is no more than a 1:1 Physical Core to Virtual CPU ratio (no oversubscribing)
    - If Hyper-Threading is enabled do not count Hyper-Threaded cores as physical cores
    - Do not oversubscribe memory or use dynamic memory allocation
    
Although Exchange technically supports up to a 2:1 physical core to vCPU ratio, a 1:1 ratio is strongly recommended for performance reasons.  Certain third party Hyper-Visors such as VMWare have their own guidance.  VMWare recommends a 1:1 ratio.  Their guidance can be found at https://www.vmware.com/files/pdf/Exchange_2013_on_VMware_Best_Practices_Guide.pdf.  For further details, please review the virtualization recommendations on TechNet at https://technet.microsoft.com/en-us/library/36184b2f-4cd9-48f8-b100-867fe4c6b579(v=exchg.150)#BKMK_Prereq.  Related specifically to VMWare, if you notice you are experiencing packet loss on your VMXNET3 adapter, you may want to review the following article from VMWare:  http://kb.vmware.com/selfservice/microsites/search.do?language=en_US&cmd=displayKC&externalId=2039495. 

"@

#System Information (WMI/CIM)
$plan = Get-WmiObject -ComputerName $Server -Class Win32_PowerPlan -Namespace root\cimv2\power -Filter "isActive='true'"
$proc = Get-WmiObject -ComputerName $Server -Class Win32_Processor
$system = Get-WmiObject -ComputerName $Server -Class Win32_ComputerSystem
$os = Get-WmiObject -ComputerName $Server -Class Win32_OperatingSystem
$pagefile = Get-WmiObject -ComputerName $Server -Class Win32_PageFileSetting
$win2008nic = Get-WmiObject -ComputerName $Server -Class Win32_NetworkAdapter | ?{$_.NetConnectionStatus -eq 2}

#Processor Information
$script:NumberOfCores = $null
$script:NumberOfLogicalProcessors = $null
$script:MegacyclesPerCore = $null
$script:ProcessorIsThrottled = $false
$script:ProcessorName = $null
$script:CurrentMegacycles = $null

#Version Information
$script:LocalServerIs2012R2OrLater = $false
$script:RemoteServerIs2012R2OrLater = $false
$script:IsExchange2010OrEarlier = $false


##################
#Helper Functions#
##################

#Output functions
function Write-Red($message)
{
    Write-Host $message -ForegroundColor Red
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Yellow($message)
{
    Write-Host $message -ForegroundColor Yellow
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Green($message)
{
    Write-Host $message -ForegroundColor Green
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Grey($message)
{
    Write-Host $message
    $message | Out-File ($OutputFullPath) -Append
}

function Write-VerboseOutput($message)
{
    Write-Verbose $message
    if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
    {
        $message | Out-File ($OutputFullPath) -Append
    }
}

function Exit-Script
{
    Write-Grey("Output file written to " + $OutputFullPath)
    Exit
}

#Check Server Role (Mailbox, CAS, Both)
function Get-ServerRole
{
    Write-VerboseOutput("Calling Get-ServerRole")

    $role = (Get-ExchangeServer $Server).ServerRole
    if($role -eq "Mailbox, ClientAccess")
    {
        [HealthChecker.ServerRole]::MultiRole
        return
    }
    elseif($role -eq "Mailbox")
    {
        [HealthChecker.ServerRole]::Mailbox
        return
    }
    elseif($role -eq "ClientAccess")
    {
        [HealthChecker.ServerRole]::ClientAccess
        return
    }
    elseif($role -eq "Edge")
    {
        [HealthChecker.ServerRole]::Edge
        return
    }  
    else
    {
        [HealthChecker.ServerRole]::None
        return
    }
}

#Get the OS version
function Get-OperatingSystemVersion
{
    Write-VerboseOutput("Calling Get-OperatingSystemVersion")

    $version = $os.Version
    if($version -eq "6.0.6000")
    {
        [HealthChecker.OSVersion]::Windows2008
        return
    }
    elseif(($version -eq "6.1.7600") -or ($version -eq "6.1.7601"))
    {
        [HealthChecker.OSVersion]::Windows2008R2
        return
    }
    elseif($version -eq "6.2.9200")
    {
        [HealthChecker.OSVersion]::Windows2012
        return
    }
    elseif($version -eq "6.3.9600")
    {
        [HealthChecker.OSVersion]::Windows2012R2
		$script:RemoteServerIs2012R2OrLater = $true
        return
    }
    else
    {
        [HealthChecker.OSVersion]::Unknown
        return
    }
}

#Get the local OS version (used to determine which powershell commands are available)
function Get-LocalOperatingSystemVersion
{
    Write-VerboseOutput("Calling Get-LocalOperatingSystemVersion")

    $version = (Get-WmiObject Win32_OperatingSystem).Version
    if($version -eq "6.0.6000")
    {
        [HealthChecker.OSVersion]::Windows2008
        return
    }
    elseif(($version -eq "6.1.7600") -or ($version -eq "6.1.7601"))
    {
        [HealthChecker.OSVersion]::Windows2008R2
        return
    }
    elseif($version -eq "6.2.9200")
    {
        [HealthChecker.OSVersion]::Windows2012
        return
    }
    elseif($version -eq "6.3.9600")
    {
        [HealthChecker.OSVersion]::Windows2012R2
		$script:LocalServerIs2012R2OrLater = $true
        return
    }
    else
    {
        [HealthChecker.OSVersion]::Unknown
        return
    }
}

#Resolve Build number to CU/RU friendly name
function Get-ExchangeUpdateName($build)
{
	switch($build)
	{
		#Exchange 2016
		{$build -eq "Version 15.1 (Build 225.16)"} {"Exchange 2016 RTM"}

		#Exchange 2013
		{$build -eq "Version 15.0 (Build 516.32)"} {"Exchange 2013 RTM"}
		{$build -eq "Version 15.0 (Build 620.29)"} {"Exchange 2013 Cumulative Update 1"}
		{$build -eq "Version 15.0 (Build 712.24)"} {"Exchange 2013 Cumulative Update 2"}
		{$build -eq "Version 15.0 (Build 775.38)"} {"Exchange 2013 Cumulative Update 3"}
		{$build -eq "Version 15.0 (Build 847.32)"} {"Exchange 2013 Service Pack 1"}
		{$build -eq "Version 15.0 (Build 913.22)"} {"Exchange 2013 Cumulative Update 5"}
		{$build -eq "Version 15.0 (Build 995.29)"} {"Exchange 2013 Cumulative Update 6"}
		{$build -eq "Version 15.0 (Build 1044.25)"} {"Exchange 2013 Cumulative Update 7"}
		{$build -eq "Version 15.0 (Build 1076.9)"} {"Exchange 2013 Cumulative Update 8"}
		{$build -eq "Version 15.0 (Build 1104.5)"} {"Exchange 2013 Cumulative Update 9"}
		{$build -like "Version 15.0 (Build 1130.7)"} {"Exchange 2013 Cumulative Update 10"}
		{$build -like "Version 15.0 (Build 1156.6)"} {"Exchange 2013 Cumulative Update 11"}
		default {""}
	}
}

#Check .NET Framework Version
#Uses registry build numbers from https://msdn.microsoft.com/en-us/library/hh925568(v=vs.110).aspx
function Get-NetFrameWorkVersion
{
    Write-VerboseOutput("Calling Get-NetFrameWorkVersion")

    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
    $RegKey= $Reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")
    [int]$NetVersionKey= $RegKey.GetValue("Release")

    if($NetVersionKey -ge 394271)
    {
        "4.6.1 or later"
        return
    }
    switch ($NetVersionKey)
    {
        {($_ -ge 378389) -and ($_ -lt 378675)} {"4.5"}
        {($_ -ge 378675) -and ($_ -lt 379893)} {"4.5.1"}
        {($_ -ge 379893) -and ($_ -lt 393297)} {"4.5.2"}
		{($_ -ge 393297) -and ($_ -lt 394271)} {"4.6"}
        default {"Unable to Determine"}
    }
}

#Write All Reported Information to the console
function Write-SystemInformationToConsole
{
    Write-VerboseOutput("Calling Write-SystemInformationToConsole")

    ###########################
    #System Information Header#
    ###########################

    Write-Green($ScriptName + " version " + $ScriptVersion)
    Write-Green("System Information Report for " + $Server + " on " + (Get-Date) + "`r`n")

    ###############################
    #OS, System, and Exchange Info#
    ###############################

    #Virtualized
    if($system.Manufacturer -like "VMWare*")
    {
        $ServerType = [HealthChecker.ServerType]::VMWare
        Write-Yellow($VirtualizationWarning)
    }
    elseif($system.Manufacturer -like "Microsoft Corporation")
    {
        $ServerType = [HealthChecker.ServerType]::HyperV
        Write-Yellow($VirtualizationWarning)
    }
    elseif($system.Manufacturer.Length -gt 0)
    {
        $ServerType = [HealthChecker.ServerType]::Physical
    }
    else
    {
        $ServerType = [HealthChecker.ServerType]::Unknown
    }

    Write-Grey("Hardware/OS/Exchange Information:")
    Write-Grey("`tHardware Type: " + $ServerType.ToString())
    if($ServerType -eq [HealthChecker.ServerType]::Physical)
    {
        Write-Grey("`tManufacturer: " + $system.Manufacturer)
        Write-Grey("`tModel: " + $system.Model)
    }

    #OS Version
    Write-Grey("`tOperating System: " + $os.Caption)

    #Exchange Version
	$version = (Get-ExchangeServer $Server).AdminDisplayVersion
    Write-Grey("`tExchange: " + $version + " " + (Get-ExchangeUpdateName($version)))
	if($version.Major -eq "14")
	{
		$script:IsExchange2010OrEarlier = $true
	}

    #ServerRole
	if($script:IsExchange2010OrEarlier)
	{
		Write-Grey("`tServer Role: " + (Get-ExchangeServer $Server).ServerRole)
	}
	else
	{
		if($ServerRole -eq [HealthChecker.ServerRole]::MultiRole)
		{
			Write-Grey("`tServer Role: " + $ServerRole.ToString())
		}
		elseif($ServerRole -eq [HealthChecker.ServerRole]::Edge)
		{
			Write-Grey("`tServer Role: " + $ServerRole.ToString())
		}
		else
		{
			Write-Yellow("`tServer Role: " + $ServerRole.ToString() + " --- Warning: Multi-Role servers are recommended")
		}
	}


    ##########
    #Pagefile#
    ##########

    Write-Grey("Pagefile Settings:")
    if($system.AutomaticManagedPagefile -eq $true)
    {
        Write-Red("`tError: System is set to automatically manage the pagefile size.  This is not recommended.")
    }
    else
    {
        if($pagefile.MaximumSize -eq 0)
        {
            Write-Grey("`tPagefile Size couldn't be detected")
        }
        #If we have more than 32GB RAM in the system, pagefile should be capped at 32GB
        elseif($system.TotalPhysicalMemory -gt 34370224128)
        {
            if($pagefile.MaximumSize -gt 32778)
            {
                Write-Yellow("`tPagefile Size: " + $pagefile.MaximumSize + " -- Pagefile should be capped at 32778 MB")
            }
            else
            {
                Write-Grey("`tPagefile Size: " + $pagefile.MaximumSize)
            }
        }
        else
        {
            Write-Grey("`tPagefile Size: " + $pagefile.MaximumSize)
        }
    }




    ################
    #.NET FrameWork#
    ################

	if(-not $script:IsExchange2010OrEarlier)
	{
		Write-Grey(".NET Framework:")
		#Report .NET Framework Version
		if($NetFrameWorkVersion -eq '4.5.2')
		{
			Write-Green("`tVersion: " + $NetFrameWorkVersion)
		}
		elseif(($NetFrameWorkVersion -eq '4.6') -or ($NetFrameWorkVersion -eq '4.6.1 or later'))
		{
			Write-Red("`tVersion: " + $NetFrameWorkVersion + " --- Error: .NET FrameWork Version 4.6 or later is not yet supported.  4.5.2 is strongly recommended.")
		}
		else
		{
			Write-Red("`tVersion: " + $NetFrameWorkVersion + " --- Error: .NET FrameWork Version 4.5.2 is strongly recommended")
		}
	}


    ################
    #Power Settings#
    ################

    Write-Grey("Power Settings:")

    #Report Power Plan
    if($plan.ElementName -eq "High performance")
    {
        Write-Green("`tPower Plan: " + $plan.ElementName)
    }
    else
    {
        Write-Red("`tPower Plan: " + $plan.ElementName + " --- Error: High performance power plan is recommended")
    }


    ##################
    #Network Settings#
    ##################

    #Report name and speed of each enabled network card, and RSS setting for 2012R2.  Windows 2008/R2/2012 do not support Get-NetworkAdapter so we have to use WMI for it.
	#Both local and remote server have to support Get-NetworkAdapter and New-CimSession for it to work.
	if($script:LocalServerIs2012R2OrLater -and $script:RemoteServerIs2012R2OrLater )
    {
		$cim = New-CimSession -ComputerName $Server
        $NetworkCards = Get-NetAdapter -CimSession $cim | ?{$_.MediaConnectionState -eq "Connected"}
        Write-Grey("NIC settings per active adapter:")
        $adaptercount = 0
        foreach($adapter in $NetworkCards)
        {
            $adaptercount++
            $RSSSettings = $adapter | Get-NetAdapterRss
            Write-Grey("`tInterface Description: " + $adapter.InterfaceDescription)
            #Get driver version and age if on physical hardware
            if($ServerType -eq [HealthChecker.ServerType]::Physical)
            {
                #warn if over 1 year old
                if((New-TimeSpan -Start (Get-Date) -End $adapter.DriverDate) -lt [int]-365)
                {
                    Write-Yellow("`t`tWarning: NIC driver is over 1 year old.  Verify you are at the latest version.")
                }
                Write-Grey("`t`tDriver Date: " + $adapter.DriverDate)
                Write-Grey("`t`tDriver Version: " + $adapter.DriverVersionString)
            }
			if(($ServerType -eq [HealthChecker.ServerType]::HyperV) -or ($ServerType -eq [HealthChecker.ServerType]::VMWare))
			{
				Write-Yellow("`t`tLink Speed: Cannot be accurately determined due to virtualized hardware")
			}
			else
			{
				Write-Grey("`t`tLink Speed: " + (($adapter.Speed)/1000000).ToString() + " Mbps")
			}
            if($RSSSettings.Enabled -eq $false)
            {
                Write-Yellow("`t`tRSS: Disabled --- Warning: Enabling RSS is recommended.")
            }
            else
            {
                Write-Green("`t`tRSS: Enabled")
            }
        }
        #if we have more than one NIC, let them know we don't need separate ones for replication and MAPI networks any more
		if(-not $script:IsExchange2010OrEarlier)
        {
			if(($adaptercount -gt 1) -and (($ServerRole -eq [HealthChecker.ServerRole]::MultiRole) -or ($ServerRole -eq [HealthChecker.ServerRole]::Mailbox)))
			{
				Write-Yellow("`t`tMultiple active network adapters detected.  Exchange 2013 may not need separate adapters for MAPI and replication traffic.  For details please refer to https://technet.microsoft.com/en-us/library/29bb0358-fc8e-4437-8feb-d2959ed0f102(v=exchg.150)#NR")
			}
		}
    }
    else
    {
        Write-Grey("NIC settings per active adapter:")
		Write-Yellow("`tMore detailed NIC settings can be detected if both the local and target server are running on Windows 2012 R2 or later.")
        $adaptercount = 0
        foreach($adapter in $win2008nic)
        {
            $adaptercount++
            Write-Grey("`tInterface Description: " + $adapter.Description)
			if(($ServerType -eq [HealthChecker.ServerType]::HyperV) -or ($ServerType -eq [HealthChecker.ServerType]::VMWare))
			{
				Write-Yellow("`tLink Speed: Cannot be accurately determined due to virtualized hardware")
			}
			else
			{
				Write-Grey("`tLink Speed: " + (($adapter.Speed)/1000000).ToString() + " Mbps")
			}
        }
		if(-not $script:IsExchange2010OrEarlier)
		{
			#if we have more than one NIC, let them know we don't need separate ones for replication and MAPI networks any more
			if(($adaptercount -gt 1) -and (($ServerRole -eq [HealthChecker.ServerRole]::MultiRole) -or ($ServerRole -eq [HealthChecker.ServerRole]::Mailbox)))
			{
				Write-Yellow("`tMultiple active network adapters detected.  Exchange 2013 may not need separate adapters for MAPI and replication traffic.`r`nhttps://technet.microsoft.com/en-us/library/29bb0358-fc8e-4437-8feb-d2959ed0f102(v=exchg.150)#NR")
			}
		}
    }



    #######################
    #Processor Information#
    #######################
    Write-Grey("Processor/Memory Information:")
    if($script:ProcessorIsThrottled -eq $true)
    {
        Write-Red("`tProcessor speed is being throttled.  Ensure the BIOS is set to allow the OS to manage power and `r`n`tthe power plan is set to `"High performance`".")
        Write-Red("`tCurrent Processor Speed: " + $script:CurrentMegacycles)
        Write-Red("`tMax Processor Speed: " + $script:MegacyclesPerCore)
    }
    Write-Grey("`tProcessor Type: " + $script:ProcessorName)
	if($script:ProcessorName.StartsWith("AMD"))
	{
		Write-Yellow("This script may incorrectly report that Hyper-Threading is enabled on certain AMD processors.  Check with the manufacturer to see if your model supports SMT.")
	}
    Write-Grey("`tPhysical Memory: " + [Math]::Round(($system.TotalPhysicalMemory)/1024/1024, 0) + " MB")
    Write-Grey("`tNumber of Processors: " + $system.NumberOfProcessors)
    Write-Grey("`tNumber of Physical Cores: " + $script:NumberOfCores)
	#recommendation by the PG is no more than 24 physical cores
	if($script:NumberOfCores -gt 24 -and (-not $script:IsExchange2010OrEarlier))
	{
		Write-Red("`tMore than 24 physical cores detected.  This is not recommended.  For details see`r`n`thttp://http://blogs.technet.com/b/exchange/archive/2015/06/19/ask-the-perf-guy-how-big-is-too-big.aspx")
	}
    Write-Grey("`tMegacycles Per Core: " + $script:MegacyclesPerCore)
    Write-Grey("`tNumber of Logical Processors: " + $script:NumberOfLogicalProcessors)
    if($script:NumberOfLogicalProcessors -gt $script:NumberOfCores)
    {
		if($script:IsExchange2010OrEarlier)
		{
			Write-Grey("`tHyper-Threading Enabled:  Yes")
		}
		else
		{
			if($script:NumberOfLogicalProcessors -gt 24)
			{
				Write-Red("`tMore than 24 logical cores detected.  Please disable Hyper-Threading.  For details see`r`n`thttp://http://blogs.technet.com/b/exchange/archive/2015/06/19/ask-the-perf-guy-how-big-is-too-big.aspx")
			}
			else
			{
				Write-Yellow("`tHyper-Threading Enabled: Yes --- Warning: Enabling Hyper-Threading is not recommended")
			}
		}
    }
    else
    {
        Write-Green("`tHyper-Threading Enabled: No")
    }

    ################
	#Service Health#
	################
	$services = Test-ServiceHealth -Server $Server | %{$_.ServicesNotRunning}
	if($services.length -gt 0)
	{
		Write-Yellow("`r`nThe following services are not running:")
		$services | %{Write-Grey($_)}
	}

	#################
	#TCP/IP Settings#
	#################
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
	$RegKey= $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\Tcpip\Parameters")
	[int]$KeepAliveValue= $RegKey.GetValue("KeepAliveTime")
	if($KeepAliveValue -eq 0)
	{
		Write-Grey("`r`nTCP/IP Settings:")
		Write-Yellow("The TCP KeepAliveTime value is not specified in the registry.  Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration.  To avoid issues, add the KeepAliveTime REG_DWORD entry under HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters and set it to a value between 900000 and 1800000 decimal.  Please note that this change will require a restart of the system.")
	}
	elseif(-not ($KeepAliveValue -ge 900000) -and ($KeepAliveValue -le 1800000))
	{
		Write-Grey("`r`nTCP/IP Settings:")
		Write-Yellow("The TCP KeepAliveTime value is not configured optimally.  It is currently set to " + $KeepAliveValue + ". This can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration.  To avoid issues, set the HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters\KeepAliveTime registry entry to a value between 15 and 30 minutes (900000 and 1800000 decimal).  Please note that this change will require a restart of the system.")
	}

	#############
	#IU\SU Check#
	#############
	if(-not $script:IsExchange2010OrEarlier)
	{
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
		$RegKey= $Reg.OpenSubKey("SOFTWARE\Microsoft\Updates\Exchange 2013\SP1")
		if($RegKey -ne $null)
		{
			$IU = $RegKey.GetSubKeyNames()
			if($IU -ne $null)
			{
				Write-Yellow("`r`nInterim Update or Security Update Detected")
			}
			foreach($key in $IU)
			{
				$IUKey = $Reg.OpenSubKey("SOFTWARE\Microsoft\Updates\Exchange 2013\SP1\" + $key)
				$IUName = $IUKey.GetValue("PackageName")
				Write-Grey($IUName)
			}
		}
	}

	##############
	#Hotfix Check#
	##############
	if(-not $script:IsExchange2010OrEarlier)
	{
		Write-Grey("`r`nHotfix Check:")
		Check-Hotfixes
	}


    ##################
    #Mailbox/DB stats#
    ##################
    if($MailboxReport)
    {
		if($script:IsExchange2010OrEarlier)
		{
			Write-Yellow("Mailbox Report not supported on versions earlier than Exchange 2013.")
			Exit
		}
        if(($ServerRole -eq [HealthChecker.ServerRole]::Mailbox) -or ($ServerRole -eq [HealthChecker.ServerRole]::MultiRole))
        {
            Write-Grey("Database and Mailbox Statistics:")
            Get-DatabaseAndMailboxStatistics
        }
        else
        {
            Write-VerboseOutput("Mailbox role not detected.  Skipping database/mailbox statistics.")
        }
    }

    ############
    #End Report#
    ############
    Write-Grey "`r`n`r`n"
}

#Get database/mailbox statistics
function Get-DatabaseAndMailboxStatistics
{
    Write-VerboseOutput("Calling Get-DatabaseAndMailboxStatistics")

    #Get a list of all active database copies on this server and total the number of mailboxes on each DB
    $MountedDBs = Get-MailboxDatabaseCopyStatus -server $Server | ?{$_.Status -eq 'Mounted'}
    Write-Grey("`tActive Databases:")
    foreach($db in $MountedDBs)
    {
        Write-Grey("`t`t" + $db.Name)
    }
    if($MountedDBs.Count -gt 0)
    {
        $MountedDBs.DatabaseName | %{Write-VerboseOutput "`tCalculating Private Mailbox Total for Active Database: $_";$TotalActivePrivateMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-Grey("`tTotal Active User Mailboxes on server: " + $TotalActivePrivateMailboxCount)
        $MountedDBs.DatabaseName | %{Write-VerboseOutput "`tCalculating Private Mailbox Total for Active Database: $_";$TotalActivePublicMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count}
        Write-Grey("`tTotal Active Public Folder Mailboxes on server: " + $TotalActivePublicMailboxCount)

        Write-Grey("`tTotal Active Mailboxes on server: " + ($TotalActivePrivateMailboxCount + $TotalActivePublicMailboxCount).ToString())
    }
    else
    {
        Write-Grey "`tNo Active Mailboxes found on server."
    }

    #Get a list of all passive database copies on this server and total the number of mailboxes on each DB.  Not on by default, must use -MailboxReport.
    $MountedDBs = Get-MailboxDatabaseCopyStatus -server $Server | ?{$_.Status -eq 'Healthy'}
    Write-Grey("`r`n`tPassive Databases:")
    foreach($db in $MountedDBs)
    {
        Write-Grey("`t`t" + $db.Name)
    }
    if($MountedDBs.Count -gt 0)
    {
        $MountedDBs.DatabaseName | %{Write-VerboseOutput "`tCalculating Private Mailbox TotalPassivePrivateMailboxCount for Passive Database: $_";$TotalPassivePrivateMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-Grey("`tTotal Passive User Mailboxes on server: " + $TotalPassivePrivateMailboxCount)
        $MountedDBs.DatabaseName | %{Write-VerboseOutput "`tCalculating Private Mailbox Total for Passive Database: $_";$TotalPassivePublicMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count}
        Write-Grey("`tTotal Passive Public Folder Mailboxes on server: " + $TotalPassivePublicMailboxCount)

        Write-Grey("`tTotal Passive Mailboxes on server: " + ($TotalPassivePrivateMailboxCount + $TotalPassivePublicMailboxCount).ToString())
    }
    else
    {
        Write-Grey("`tNo Passive Mailboxes found on server.")
    }
}

#Check for hotfixes that are recommended to prevent known high CPU/performance issues in Exchange 2013/2016
function Check-Hotfixes
{
	$2008HotfixList = @("")
	$2008R2HotfixList = @("KB3004383")
	$2012HotfixList = @("")
	$2012R2HotfixList = @("KB3041832")
	$hotfixesneeded = $false

	switch(Get-OperatingSystemVersion)
	{
		#Windows 2008R2
		{ ($_ -eq [HealthChecker.OSVersion]::Windows2008) } 
			{ 
				if($2008HotfixList -ne $null)
				{
					foreach($hotfix in $2008HotfixList)
					{
						if((Get-HotFix -Id $hotfix -ErrorAction SilentlyContinue) -eq $null)
						{
							$hotfixesneeded = $true
							Write-Yellow("Hotfix " + $hotfix + " is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues.")
						}
					}
				}
			}
		
		#Windows 2008R2
		{ ($_ -eq [HealthChecker.OSVersion]::Windows2008R2) } 
			{ 
				if($2008R2HotfixList -ne $null)
				{
					foreach($hotfix in $2008R2HotfixList)
					{
						if((Get-HotFix -Id $hotfix -ErrorAction SilentlyContinue) -eq $null)
						{
							$hotfixesneeded = $true
							Write-Yellow("Hotfix " + $hotfix + " is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues.")
						}
					}
				}
			}

		#Windows 2012
		{ ($_ -eq [HealthChecker.OSVersion]::Windows2012) } 
			{ 
				if($2012HotfixList -ne $null)
				{
					foreach($hotfix in $2012HotfixList)
					{
						if((Get-HotFix -Id $hotfix -ErrorAction SilentlyContinue) -eq $null)
						{
							$hotfixesneeded = $true
							Write-Yellow("Hotfix " + $hotfix + " is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues.")
						}
					}
				}
			}

		#Windows 2012R2
		{ ($_ -eq [HealthChecker.OSVersion]::Windows2012R2) } 
			{
				if($2012R2HotfixList -ne $null)
				{ 
					foreach($hotfix in $2012R2HotfixList)
					{
						if((Get-HotFix -Id $hotfix -ErrorAction SilentlyContinue) -eq $null)
						{
							$hotfixesneeded = $true
							Write-Yellow("Hotfix " + $hotfix + " is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues.")
						}
					}
				}
			}
	}

	if(-not $hotfixesneeded)
	{
		Write-Grey("Hotfix check complete.  No action required.")
	}
}

#Check CAS load balancing
function Get-CASLoadBalancingReport
{
    Write-VerboseOutput("Get-CASLoadBalancingReport")

	if((Get-ExchangeServer $Server).AdminDisplayVersion.Major -eq "14")
	{
		Write-Yellow("-LoadBalancingReport not supported on versions earlier than Exchange 2013")
		Exit
	}

    #Connection and requests per server and client type values
    $CASConnectionStats = @{}
    $TotalCASConnectionCount = 0
    $AutoDStats = @{}
    $TotalAutoDRequests = 0
    $EWSStats = @{}
    $TotalEWSRequests = 0
    $MapiHttpStats = @{}
    $TotalMapiHttpRequests = 0
    $EASStats = @{}
    $TotalEASRequests = 0
    $OWAStats = @{}
    $TotalOWARequests = 0
    $RpcHttpStats = @{}
    $TotalRpcHttpRequests = 0

    #List of CAS servers to operate against.  Default is all 2013 CAS.  Alternatively you can use the
    #-CasServerList switch to specify the servers, separated by comma.
    $CASServers = @()
    if($CasServerList -ne $null)
    {
        foreach($cas in $CasServerList)
        {
            $CASServers += (Get-ExchangeServer $cas)
        }
    }
    else
    {
        $CASServers = Get-ExchangeServer | ?{($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15")}
    }

    #Pull connection and request stats from perfmon for each CAS
    foreach($cas in $CASServers)
    {
        #Total connections
        $TotalConnectionCount = (Get-Counter ("\\" + $cas.Name + "\Web Service(Default Web Site)\Current Connections")).CounterSamples.CookedValue
        $CASConnectionStats.Add($cas.Name, $TotalConnectionCount)
        $TotalCASConnectionCount += $TotalConnectionCount

        #AutoD requests
        $AutoDRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Autodiscover)\Requests Executing")).CounterSamples.CookedValue
        $AutoDStats.Add($cas.Name, $AutoDRequestCount)
        $TotalAutoDRequests += $AutoDRequestCount

        #EWS requests
        $EWSRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_EWS)\Requests Executing")).CounterSamples.CookedValue
        $EWSStats.Add($cas.Name, $EWSRequestCount)
        $TotalEWSRequests += $EWSRequestCount

        #MapiHttp requests
        $MapiHttpRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_mapi)\Requests Executing")).CounterSamples.CookedValue
        $MapiHttpStats.Add($cas.Name, $MapiHttpRequestCount)
        $TotalMapiHttpRequests += $MapiHttpRequestCount

        #EAS requests
        $EASRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests Executing")).CounterSamples.CookedValue
        $EASStats.Add($cas.Name, $EASRequestCount)
        $TotalEASRequests += $EASRequestCount

        #OWA requests
        $OWARequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_owa)\Requests Executing")).CounterSamples.CookedValue
        $OWAStats.Add($cas.Name, $OWARequestCount)
        $TotalOWARequests += $OWARequestCount

        #RPCHTTP requests
        $RpcHttpRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Rpc)\Requests Executing")).CounterSamples.CookedValue
        $RpcHttpStats.Add($cas.Name, $RpcHttpRequestCount)
        $TotalRpcHttpRequests += $RpcHttpRequestCount
    }

    #Report the results for connection count
    Write-Grey("")
    Write-Grey("Connection Load Distribution Per Server")
    Write-Grey("Total Connections: " + $TotalCASConnectionCount)
    #Calculate percentage of connection load
    $CASConnectionStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
    Write-Grey($_.Key + ": " + $_.Value + " Connections = " + [math]::Round((([int]$_.Value/$TotalCASConnectionCount)*100)) + "% Distribution")
    }

    #Same for each client type.  These are request numbers not connection numbers.
    #AutoD
    if($TotalAutoDRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current AutoDiscover Requests Per Server")
        Write-Grey("Total Requests: " + $TotalAutoDRequests)
        $AutoDStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalAutoDRequests)*100)) + "% Distribution")
        }
    }

    #EWS
    if($TotalEWSRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current EWS Requests Per Server")
        Write-Grey("Total Requests: " + $TotalEWSRequests)
        $EWSStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalEWSRequests)*100)) + "% Distribution")
        }
    }

    #MapiHttp
    if($TotalMapiHttpRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current MapiHttp Requests Per Server")
        Write-Grey("Total Requests: " + $TotalMapiHttpRequests)
        $MapiHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalMapiHttpRequests)*100)) + "% Distribution")
        }
    }

    #EAS
    if($TotalEASRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current EAS Requests Per Server")
        Write-Grey("Total Requests: " + $TotalEASRequests)
        $EASStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalEASRequests)*100)) + "% Distribution")
        }
    }

    #OWA
    if($TotalOWARequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current OWA Requests Per Server")
        Write-Grey("Total Requests: " + $TotalOWARequests)
        $OWAStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalOWARequests)*100)) + "% Distribution")
        }
    }

    #RpcHttp
    if($TotalRpcHttpRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current RpcHttp Requests Per Server")
        Write-Grey("Total Requests: " + $TotalRpcHttpRequests)
        $RpcHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalRpcHttpRequests)*100)) + "% Distribution")
        }
    }

    Write-Grey("")
}


#On multi proc boxes, WMI reports number of cores and megacycles per core as an array value for each proc such as @(8,8,8,8) instead of 32.
#It can also put the results of Get-WmiObject Win32_Processor into an array of Win32_Processor objects depending on the hardware setup.  
#Need to normalize these numbers to avoid errors.
function Normalize-ProcessorInfo
{
    Write-VerboseOutput("Calling Normalize-ProcessorInfo")

    #Handle single and multi proc machines slightly differently due to the way Win32_Processor returns the data
    if($system.NumberOfProcessors -gt 1)
    {
        #Total cores in all processors
        foreach($processor in $proc)
        {
            $coresum += $processor.NumberOfCores
            $logicalsum += $processor.NumberOfLogicalProcessors
            if($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed)
            {
                $script:CurrentMegacycles = $processor.CurrentClockSpeed
                $script:ProcessorIsThrottled = $true
            }
        }
        $script:NumberOfCores = $coresum
        $script:NumberOfLogicalProcessors = $logicalsum
     
        #all processors should be the same speed and type so take the description and Max Speed of the first processor
        $script:ProcessorName = $proc[0].Name
        $script:MegacyclesPerCore = $proc[0].MaxClockSpeed
    }
    else #single processor machine
    {
        $script:NumberOfCores = $proc.NumberOfCores
        $script:NumberOfLogicalProcessors = $proc.NumberOfLogicalProcessors
        $script:MegacyclesPerCore = $proc.MaxClockSpeed
        $script:ProcessorName = $proc.Name
        if($proc.CurrentClockSpeed -lt $proc.MaxClockSpeed)
        {
            $script:CurrentMegacycles = $proc.CurrentClockSpeed
            $script:ProcessorIsThrottled = $true
        }
    }

    #We need processor count, cores, logical processors, and megacycles to continue.  If one of these is missing, exit the script.
    if(($script:NumberOfCores -eq $null) -or ($script:NumberOfLogicalProcessors -eq $null) -or ($script:MegacyclesPerCore -eq $null))
    {
        Write-Red("Processor information could not be read.  Exiting script.")
        Exit
    }
}

#Main script execution
Write-VerboseOutput("Calling Main Script Execution")

if(-not (Test-Path $OutputFilePath))
{
    Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
    Exit
}

#Load Balancing Report
if($LoadBalancingReport)
{
    Write-Green($ScriptName + " version " + $ScriptVersion)
    Write-Green("Client Access Load Balancing Report on " + (Get-Date) + "`r`n")
    Get-CASLoadBalancingReport
    Exit-Script
}

#Normalize processor values
Normalize-ProcessorInfo

#Populate server role, OS version, and .NET framework version info
$ServerRole = Get-ServerRole
$OSVersion = Get-OperatingSystemVersion
$LocalOSVersion = Get-LocalOperatingSystemVersion
$NetFrameWorkVersion = Get-NetFrameWorkVersion

#Display system information and recommendations check results
Write-SystemInformationToConsole

#Finish
Exit-Script



