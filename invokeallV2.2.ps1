Function Invoke-All{

    <#
    .SYNOPSIS
        There were many instances where we pipe cmdlets and wait for the output for hours or days, 
        the commands simply takes so long just because of too many objects it has to process sequentially.
        This function uses runspaces to achieve multi-threading to run powershell commands.
        I have made the function very generic, easy-to-use and lightweight as possible.

        TODO:
            Handle Executables.
            Extend to capture the Error and Verbose outputs from the Job streams

    .DESCRIPTION
        Invoke-All is the function exported from the Module. It is a wrapper like function which takes input from Pipeline or using explict parameter and process the command consequently using runspaces.
        I have mainly focussed on Exchange On-Perm and Online Remote powershell (RPS) while developing the script. 
        Though I have only tested it with Exchange Remote sessions and Snap-ins, it can be used on any powershell cmdlet or function.

        Exchange remote powershell uses Implict remoting, the Proxyfunctions used by the Remote session doesnt accept value from pipeline.
        Therefore the Script requires you to declare all the Parameters that you want to use in the command. See examples on how to do it.


    .NOTES
        
        Version        : 2.2
        Author         : Santhosh Sethumadhavan (santhse@microsoft.com)
        Prerequisite   : Requires Powershell V3 or Higher.

        V 1.1 - Try/Finally block fixes on the End block
        v 2.0 - Some optimizations on the finally blocks.
                Added support for External scripts.
        v 2.1 - Added Batching support
        v 2.2 - Optimized Job collection

    .PARAMETER ScriptBlock
        Scriptblock to execute in parallel.
        Usually it is the 2nd Pipeline block, wrap your command as shown in the below examples and it should work fine.
        You cannot use alias or external scripts. If you are using a function from a custom script, please make sure it is an Advance function or with Param blocks defined properly.

    .PARAMETER InputObject
        Run script against these specified objects. Takes input from Pipeline or when specified expilictly.

    .PARAMETER MaxThreads
        Number of threads to be executed parallely, by default it creates one thread per CPU

    .PARAMETER RPS
        Use this switch if you want to run the command using Remote Powershell.
        By deafult, script autodetects it, but can be passed as a parameter as well

    .PARAMETER Force
        By default this script does error checking for the first instance. If the first job fails, likely all jobs would, so just bail out. 
        Use this parameter if you know the parameters passed are correct and also useful when you are Scheduling the script and should not be prompted

    .PARAMETER WaitTimeOut
        When Force switch is not mentioned, the script waits for the first job to complete for Error checking.
    
    .PARAMETER ModulestoLoad
        Powershell Module names that needs to be loaded in to the runspace that is required to exectue the commands used in the External script (PS1).
        This parameter is not required if you are using a command from any powershell Module

    .PARAMETER SnapinstoLoad
        Powershell Snapin names that needs to be loaded in to the runspace that is required to exectue the commands used in the External script (PS1).
        This parameter is not required if you are using a command from any powershell Module
      
    .PARAMETER BatchSize
        By default the function uses BatchSize of 100. This parameter is to limit the number of jobs that are to be Queued for running at a time and to process it in batches.
        For Example, If BatchSize is 20
        Queue 40 Jobs (BatchSize * 2)
        Wait until 20 jobs completes
        Queue the next batch of 20

    .PARAMETER PauseInMsec
        Delay to be induced between each batch. Before Queueing the subsequent batches, the script pauses for MilliSeconds mentioned.
        BatchSize and PauseInMsec can be used if Multithreading is overloading the Source or destination.
        It's very useful when running jobs against Exchange online, there are no magic numbers for this parameter, use wisely.

    .PARAMETER Quiet
        Doesnt display the progress bar. Can be used when scheduling the script.
        Using Quiet mode also speeds up the script as displaying progress bar is not required.


    .Example
        
        Get-Mailbox -database 'db1' | invoke-all {Get-MailboxFolderStatistics -FolderScope "inbox" -IncludeOldestAndNewestItems } -Force |?{$_.ItemsInFolder -gt 0} | Select Identity, Itemsinfolder, Foldersize, NewestItemReceivedDate
        
        Above Command was ran from powershell with Module or snap-in loaded.
        Actual command:
        Get-Mailbox -database 'db1' | Get-MailboxFolderStatistics -FolderScope "inbox" -IncludeOldestAndNewestItems |?{$_.ItemsInFolder -gt 0} | Select Identity, Itemsinfolder, Foldersize, NewestItemReceivedDate
        
    .Example
        
        Get-Mailbox -database 'db1' | invoke-all {Get-MailboxFolderStatistics -Identity $_.name -FolderScope "inbox" -IncludeOldestAndNewestItems } -Force |?{$_.ItemsInFolder -gt 0}
        
        Above command was ran from RPS (Exchange mangement Shell or Cloud PS session)
        Actual command:
        Get-Mailbox -database 'db1' | Get-MailboxFolderStatistics -FolderScope "inbox" -IncludeOldestAndNewestItems |?{$_.ItemsInFolder -gt 0}
       
        Note: When running from RPS, we need to specifiy all the parameters to the cmdlet for this function to work

    .Example
        
        Get-AzureRmVM -ResourceGroupName 'MyAzureRG' | Invoke-all { Start-AzureRmvm -ResourceGroupName $_.ResourceGroupname -Name $_.Name } -Force
        Command ran from Azure module imported PS console. This command Starts all VMs in Parallel.

        Actual Command:
        Get-AzureRmVM -ResourceGroupName 'MyAzureRG' | foreach { Start-AzureRmvm -Name $_.name -ResourceGroupName $_.ResourceGroupname }

    .Example

        Get-Mailbox -ResultSize 100 | Invoke-All {.\GetPWDsetUsrs.ps1 -Name $_.Alias} -ModulestoLoad Activedirectory
        Above command is an example on how to use the function with External scripts.

        ---------------------------------------GetPWDsetUsrs.ps1-----------------------------------------------
        param(
	    [String]$Name = ''
        )

            $90_Days = (Get-Date).adddays(-90)
            return Get-ADUser -filter {(mailnickname -eq $Name) -and (passwordlastset -le $90_days)} -properties PasswordExpired, PasswordLastSet, PasswordNeverExpires
        ---------------------------------------------------------------------------------------------------------

    .Example
        Get-ADComputer -Filter {name -like "*fs*"} | Invoke-All { C:\scripts\get-uptime.ps1 -ComputerName $_.DnsHostname}

    .Example
        $MBX | Invoke-All { Get-MobileDeviceStatistics -Mailbox $_.userprincipalname } -PauseInMsec 500 | `
        Select-Object @{Name="DisplayName";Expression={$input.Displayname}},Status,DeviceOS,DeviceModel,LastSuccessSync,FirstSyncTime | `
        Export-Csv c:\temp\devices.csv –Append

        Collects Mobile device statistics in batches of 100 at a time and Sleeps for 500 MilliSeconds between batches.
        This is very useful if you are using the script on a remote powershell and getting throttled, especially with Exchange online.
        

    #>




[cmdletbinding(SupportsShouldProcess = $True,DefaultParameterSetName='ScriptBlock')]
Param (   
        [Parameter(Mandatory=$True,position=0,ParameterSetName='ScriptBlock')]
            [System.Management.Automation.ScriptBlock]$ScriptBlock,
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName='ScriptBlock')]
        $InputObject,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [int]$MaxThreads = ((Get-WmiObject Win32_Processor) | Measure-Object -Sum -Property NumberOfLogicalProcessors).Sum,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [SWITCH]$RPS = $false,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [SWITCH]$Force = $false,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [INT]$WaitTimeOut = 30,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [ARRAY]$ModulestoLoad = @(),
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [ARRAY]$SnapinstoLoad = @(),
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [INT]$BatchSize = 100,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [INT]$PauseInMsec = 0,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='ScriptBlock')]
        [SWITCH]$Quiet
)


Begin{

#Discover the command that is being ran and prepare the Runspacepool objects

    If($host.Version.Major -lt [INT]'3'){
    
        Write-Host "This Script requires Powershell version 3 or greater." -ForegroundColor Red
        Break
    
    }

#region begin init

    [String]$Global:Command = ''
    [String]$ProxyCommand = ''
    [BOOL]$SupportsValfromPipeLine = $false
    
    $Commandinfo = $NULL
    $Commandtype = New-Object System.Management.Automation.CommandTypes
    $runspacepool = $NULL
    $Jobs = @{} #using Hashtable for performance
    [int]$i = 0
    [int]$JobCounter = $BatchSize * 2
    [int]$TotItems = $Input.Count
    [int]$script:jobsCollected = 0
    $Code = $NULL
    
    #[System.Management.Automation.Runspaces.WSManConnectionInfo]$Coninfo = $NULL
    $MetaCommand = ''
    $MetaData = ''
    $Coninfo = New-Object System.Management.Automation.Runspaces.WSManConnectionInfo

    $Timer = [system.diagnostics.stopwatch]::StartNew()


#endregion init

#region begin Functions

    #logging function
    function Write-Log
    {
    [CmdletBinding()] 
        Param 
        (
            [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()] 
            [String]$Message,

            [Parameter(Mandatory=$false)]
            [ValidateScript({ Test-Path "$_" })]
            [string]$LogPath= "$($PSScriptRoot)"
        
        )

        $LogPath = $LogPath.TrimEnd("\")
        if(Test-Path "$LogPath\InvokeAll.log"){
            $LogFile = "$LogPath\InvokeAll.log"
        
        }Else{
            $LogFile = New-Item -Name 'InvokeAll.log' -Path $LogPath -Force -ItemType File
            Write-Verbose "Created New log file $LogFile"
        }

    
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$FormattedDate] $Message" | Out-File -FilePath $LogFile -Append    

    }

    #Load the modules that are specfied as parameter by the user.
    Function LoadModules(){
    
        $Cmderr = $NULL
        foreach($Snapin in $SnapinstoLoad){
                   
            try{
            Get-PSSnapin $Snapin -ErrorAction Stop -Registered -Verbose:$false | Foreach{
            [void]$Sessionstate.ImportPSSnapIn($_.Name ,[ref]$null)
            Write-Verbose "Added Snapin $($_.name)"
            Write-Log "Added Snapin $($_.name)"
            }
                   
                   
            }Catch{
                Write-Error "Unable to Load Snapin $Snapin"
                Write-Log "Unable to Load Snapin $Snapin"
            }
                
        }
        foreach($Module in $ModulestoLoad){

            try{
                    
            Get-Module $Module -ListAvailable -All -ErrorAction Stop -Verbose:$false | Foreach{
                [VOID]$sessionstate.ImportPSModule($_.Name)
                Write-Verbose "Added Module $($_.name)"
                Write-Log "Added Module $($_.name)"
            }
            }Catch{
                    Write-Error "Unable to Load Module $Module"
                    Write-Log "Unable to Load Module $Module"
            }
        }
    
    }

    #Collect the Jobs that are completed. The function will return when the Jobs collected are greater or equal to the Batchsize

    Function CollectJobs{
    Param([hashtable]$Jobs,[int]$BSize, [ref]$Jobscollected)

        [int]$CurCollection = 0
        $JobsInProgress = @()


        do{

            if(($Jobs.Values.Handle |?{$_ -ne $NULL}).count -le 0){ #Break if all jobs are completed
                Return
            }    
            
            if(($Jobs.Values.Handle.IsCompleted |?{$_ -eq $True}).Count -le 0){#If there are no completed jobs, wait for atleast one using Event, else loop the Circles simply consuming resources
                
                #If 80%, Go back and Invoke more jobs than waiting for all jobs to complete, may be there are a few jobs in this batch that will run for longer time
                if((($CurCollection / $BSize)*100) -gt 80){ 
                    Return 
                }

                If(-Not $Quiet) {Write-Progress -id 2 -Activity "Running Jobs" `
                    -Status "$JobNum / $TotalObjects Jobs Invoked. $(@($Jobs.Values.GetEnumerator() |?{$_.Thread.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Running}).Count) Jobs are in Running State..Waiting for atleast one Job to complete" `
                    -PercentComplete $(($JobNum / $TotalObjects)*100)
                   
                }
                
                #WaitAny has 64 Handle limitation. limit the wait on first 60 handles
            
                $JobsInProgress = $Jobs.Values.GetEnumerator() |?{$_.Thread.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Running} | Select -First 60

                if($JobsInProgress){
                    Write-Verbose "Waiting on the Handle for job completion"
                    $h = [System.Threading.WaitHandle]::WaitAny($JobsInProgress.Handle.AsyncWaitHandle)    
                }
            }            
            
<#
            If(-not $Quiet){
            Write-Progress -id 2 -Activity "Running Jobs" `
            -Status "$JobNum / $TotalObjects Jobs Invoked."`
            -PercentComplete $(($JobNum / $TotalObjects)*100)
            #$(($Jobs.Values.thread.InvocationStateInfo.State |?{$_ -eq [System.Management.Automation.PSInvocationState]::Running}).count) Jobs are in Running State..." `
            }
#>
   	        #Collect All jobs that are completed, it could be greater than the Batchsize, but its ok
            ForEach ($Job in $($Jobs.Values.GetEnumerator() | Where-Object {$_.Handle.IsCompleted -eq $True})){
                If(-not $Quiet){
                Write-Progress `
	            -id 22  -ParentId 2 `
                -Activity "Collecting Jobs results that are completed... BatchSize: $BSize " `
		        -PercentComplete ($Jobscollected.value / $Jobs.Count * 100) `
		        -Status "$($Jobscollected.Value) / $($Jobs.Count)"
                }

                $CurCollection++;$Jobscollected.value++
                
                try{

                $Job.Thread.EndInvoke($Job.Handle) #Collect the result of the completed Job
            
                }Catch{

                Write-Error "Error on Thread EndInvoke : $_"
                Write-Log "Error on Thread EndInvoke : $_"
                Write-Host "It was processing Object $($Job.Object) . Job ID: $($Job.ID)" -ForegroundColor Yellow
                Write-Log "It was processing Object $($Job.Object) . Job ID: $($Job.ID)"

                }finally{

                if ($Job.Thread.HadErrors) {
                
                    $Job.Thread.Streams.Error.ReadAll() | % { 
                        Write-Error "The pipeline had error $_ "
                        Write-Host "It was processing Object $($Job.Object) . Job ID: $($Job.ID)" -ForegroundColor Yellow
                        Write-Log "The pipeline had error $_ "
                        Write-Log "It was processing Object $($Job.Object) . Job ID: $($Job.ID)"
                        }
                }

		        $Job.Thread.Dispose()
		        $Job.Thread = $Null
		        $Job.Handle = $Null
                
                }
            }




        }until($CurCollection -ge $BSize) #Collect until the BatchSize is reached

    
    }

#endregion Functions


Write-Log "============================ Starting to Execute - Invoke-all 2.2 =============================================="
Write-Log "$($myinvocation.Line)"

#region begin CommandDiscovery
#Find the command being used, we need this to create the Proxy command which is used to identify the parameters and its values used.

    #Collect the command details
    try{
    $CommandAsts = $scriptblock.Ast.FindAll({$args[0] -is [System.Management.Automation.Language.CommandAst]} , $true)
    $Elements = $CommandAsts.GetEnumerator().commandelements

    $Element = $elements[0]

    Write-Verbose "Got command $($Element.Value) to process"
    Write-Log "Got command $($Element.Value) to process"
    
    switch($element.gettype().name){
    

        StringConstantExpressionAst { #Anything that has single quote and double Quote is also of this type
            if($element.StringConstantType -eq 'Bareword'){#then it is a command
                $CommandInfo = Get-Command $element.value -ErrorAction Stop -ErrorVariable Cmderr
                $Commandtype = $CommandInfo.CommandType
                switch($Commandtype){
                    Cmdlet         { $Global:Command = $element.Value}
                    Function       { $Global:Command =  $element.Value}
                    ExternalScript { $Global:Command = $element.Value}
            
                }
            }
            #else{throw "Not able to recoginize the command"}             
        }

    }
    }Catch{
        Write-Error "Error processing command: $Cmderr"
        Write-Log "Error processing command: $Cmderr"
        Write-Host "Unable to Process, Please check the command that is passed on the script block" -ForegroundColor Yellow
        break
    
    }

#endregion CommandDiscovery

#region begin CreateProxyCmd
#If we have found what command is being ran, create the Proxy command.
#Proxy command is used to find the Parameters that are used on the actual command.
#This way, we dont have to copy all the local variables and other stuff to the each Runspace.


    If($Global:Command){ #create the proxy command

        $MetaData = New-Object System.Management.Automation.CommandMetaData ($Commandinfo)
        if($MetaData.Parameters.Count -le 0 -and (-not $Force)){
            Write-Error "$Global:Command doesnt use any parameters, the Input parameter from Pipeline cannot be bound correctly to this command"
            Write-Log "$Global:Command doesnt use any parameters, the Input parameter from Pipeline cannot be bound correctly to this command"
            Write-Host "If it is a custom script, make sure the Param blocks are defined correctly" -ForegroundColor Yellow
            Break
        
        }

        $PScript = [System.Management.Automation.ProxyCommand]::Create($MetaData)

        $PScript = [scriptblock]::Create($PScript) 

        $Paramblock = $PScript.ast.ParamBlock.ToString()

        #$strfunc = "[CmdletBinding()] `n"
        
        $strfunc = $strfunc + $Paramblock
        #$Strfunc += "`n`$PSBoundParameters.Add('`$args', `$args)"
        $strfunc += "`n `$ParamsPassed = `$PSBoundParameters `n"
        $strfunc += "return `$ParamsPassed"

        if(($strfunc.ToLower()).contains('valuefrompipeline')){ 
            $SupportsValfromPipeLine = $True
            Write-Verbose "This command supports ValueFromPipeline"
            Write-Log "This command supports ValueFromPipeline"
        }
        
        if($Commandtype -eq 'ExternalScript'){
            $ProxyCommand = ($Global:Command.Split('\')[-1]).replace(".ps1","Proxy.ps1")
            $Code = [ScriptBlock]::Create($(Get-Content $MetaData.Name | Out-String))
        }else{
            $ProxyCommand = "$Global:Command" + "Proxy"
        }

        try{

        if(Get-Command -CommandType function |?{$_.name -eq "$ProxyCommand"}){
            Remove-Item function:\$ProxyCommand -Confirm:$false #Remove if there is a duplicate, it is there from a previous failure
        }
        New-Item -Path function:global:$ProxyCommand -Value $strfunc -ErrorAction Stop -ErrorVariable Cmderr | Out-Null
        
        }catch{
        Write-Error "Unable to create the Proxy command, Error : $cmderr"
        Write-Log "Unable to create the Proxy command, Error : $cmderr"
        Break
        }

        Write-Verbose "Created the Proxy command $ProxyCommand"
        

    }else{

        Write-Error "Sorry, This script is not capable of handling the command or alias you passed"
        Write-Log "Sorry, This script is not capable of handling the command or alias you passed"
        Break

    }

#endregion CreateProxyCmd

#region begin PrepRunspace
#Possiblites are RemotePowershell or Local. If it is local we will need to identify the module that is required to run the command and load it.


    #Find if the command is from (Exchange) Remote powershell
    if($Commandinfo.Module){ 

        if(-not $RPS -and ((Get-Module $CommandInfo.Module).Description.Contains("Implicit remoting"))){ #Detect RPS even though it is not mentioned explicity
    
            Write-Verbose "Setting RPS to True"
            Write-Log "Setting RPS to True"
            $RPS = $True
    
        }
    }
    
    #Create the sessionstate object based on the command discovery

    if($RPS){
        $RemoteSession = Get-PSSession |?{$_.state -eq 'opened' -and $_.configurationname -eq 'Microsoft.Exchange'} | Select -First 1
        If(-not $RemoteSession){
            Write-Host "Unable to find the session configuration, please reconnect the session and try again" -ForegroundColor Red
            Write-Log "Unable to find the session configuration, please reconnect the session and try again"
            break
        }
        
        #if a valid RPS is found, copy the connectionInfo to use it on the runspaces

        $Coninfo = $RemoteSession.Runspace.ConnectionInfo.copy()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $Coninfo, $Host)
    }Else{
        $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        #Find module or PSsnapin to load to the session state
        if($Commandinfo.ModuleName){
            if((Get-Module $Commandinfo.ModuleName -ErrorAction SilentlyContinue)){
                [VOID]$sessionstate.ImportPSModule( $Commandinfo.ModuleName )
                Write-Verbose "Imported Module $($Commandinfo.ModuleName)"
                Write-Log "Imported Module $($Commandinfo.ModuleName)"
            }else{
                if((Get-PSSnapin $Commandinfo.ModuleName -ErrorAction SilentlyContinue)){
                    [void]$Sessionstate.ImportPSSnapIn($Commandinfo.ModuleName,[ref]$null)
                    Write-Verbose "Imported PSSnapin $($Commandinfo.ModuleName)"
                    Write-Log "Imported PSSnapin $($Commandinfo.ModuleName)"
                }
            }
        }
        else{ #If module is not found, its likely a local function or script
            
            #Load the modules that are specified by the user
            LoadModules

            #Check if it is a custom function from a local script and load the function to sessionstate
            if((Get-item Function:\$Global:Command -ErrorAction SilentlyContinue).ScriptBlock.File){
                Write-Verbose "The command $Global:Command is a custom function from file $((Get-item Function:\$Global:Command -ErrorAction SilentlyContinue).ScriptBlock.File)"
                Write-Log "The command $Global:Command is a custom function from file $((Get-item Function:\$Global:Command -ErrorAction SilentlyContinue).ScriptBlock.File)"
                $Definition = Get-Content Function:\$Global:Command
                $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $Global:Command, $Definition
                [VOID]$sessionstate.Commands.Add($SessionStateFunction)

            }elseif($Commandtype -eq 'ExternalScript'){ #We have already loaded the modules and Snapins that are specified by user
                Write-Verbose "The command passed is an External Script"
                Write-Log "The command passed is an External Script"
                
            }else{
                Write-Verbose "Unable to find the Module or snap-in to load"
            }
        }
            
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $sessionstate, $Host)
    }
    $runspacepool.Open() 
    
    
#endregion PrepRunspace

    $Timer.Start() #Trace time took on Process and End Blocks

} #End Begin

Process{ #This Block runs for each object

#Foreach object discover the Parameters and add it to the Job Object


    $error.clear()    
    $tempScriptBlock = ''
    $paramused = $NULL
    
        
    $i++

#region begin ParamDiscovery    
#Not all cmdlets and functions takes valuefromPipeline and some commands require mandatory parameters, its easier for user to specify explicitly.
#try using the Pipelineobject and if it fails try without it
    
    $tempScriptBlock = ($ScriptBlock.ToString()).Replace($Global:Command,$ProxyCommand)
    $tempScriptBlock = $tempScriptBlock.Replace('$_', '$inputobject')
    $ScriptBlockBKP = $tempScriptBlock

    if($SupportsValfromPipeLine){
        $tempScriptBlock = "`$inputobject | " + $tempScriptBlock
    }

    $tempscriptblock = [Scriptblock]::Create($tempScriptBlock)

    #lets try with the Pipelineobject and without

    try{
    $paramused = Invoke-Command -ScriptBlock $tempScriptBlock -ErrorAction SilentlyContinue -ErrorVariable Cmderr
    
    if ($Cmderr -and $Cmderr[-1].Exception.ErrorId -eq 'InputObjectNotBound'){

        Write-Host "Encountered error, but it's ok, retrying without the Pipeline Inputobject" -ForegroundColor Yellow
        Write-Log "Encountered error, but it's ok, retrying without the Pipeline Inputobject"

    #Retry without the Pipeline object
        $tempScriptBlock = [Scriptblock]::Create($ScriptBlockBKP)
        $paramused = Invoke-Command -ScriptBlock $tempScriptBlock -ErrorAction SilentlyContinue -ErrorVariable Cmderr

        If($Cmderr){

            Write-Error "Unable to execute the proxy command $($ProxyCommand) . Please verify if mandatory parameters are mentioned, if Command accepts ValueFromPipeline, do not explicity mention it as a Parameter."
            Write-Log "Unable to execute the proxy command $($ProxyCommand)"
            throw $Cmderr
        }

        $SupportsValfromPipeLine = $false
    }
 
#endregion ParamDiscovery

#region begin ErrChk    
    $Handle = $NULL
    #Check if we are able to execute the command for the first object. If it fails, likely all threads would, so quit right there

    if($i -eq 1 -and (-not $Force)){

        $Powershell = [powershell]::Create()
    
        if($Commandtype -eq 'ExternalScript'){
            [VOID]$Powershell.AddScript($Code)
        }else{
            [VOID]$Powershell.AddCommand($Global:Command)
        }
        foreach($item in $paramused.GetEnumerator()){

            $Powershell.AddParameter($item.Key,$item.value) | Out-Null
        }

        $Powershell.RunspacePool = $RunspacePool
        
        $Handle = $Powershell.BeginInvoke() #Invoke the first instance 
        
        if($Powershell.InvocationStateInfo.State -ne [System.Management.Automation.PSInvocationState]::Completed){
            Write-Verbose "Running error check on the first instance"
            Write-Log "Running error check on the first instance"
        
            Register-ObjectEvent -InputObject $Powershell -EventName InvocationStateChanged -SourceIdentifier PSInvocationStateChanged
            

            Write-Verbose "Waiting for the Invocation state change, waits for $WaitTimeOut Seconds"
            Write-Verbose "Current State : $($Powershell.InvocationStateInfo.State)"
            
            if($Powershell.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Running){
                Wait-Event -SourceIdentifier PSInvocationStateChanged -Timeout $WaitTimeOut | Out-Null
            }
            #Still running?
            if($Powershell.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Running){

                if($PSCmdlet.Shouldcontinue($Global:Command,"First Instance of the command is still running. `
                Do you want to ABORT the operation ? `
                Selecting NO will continue to wait for the first instance to Complete Indefinitely before Multi-Threading rest of the instances")){
                    
                    #Cleanup and Quit
                    Throw "Aborting the operation"                
                                   
                }

                Write-Verbose "Waiting for the first instance to complete"
                #wait forever for the first instance to complete
                Wait-Event -SourceIdentifier PSInvocationStateChanged | Out-Null

            }

            Write-Verbose "Done waiting on the first Instance"
        }

       if($Powershell.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Failed){

            Write-Error "Error invoking powershell thread. Reason: $($Powershell.InvocationStateInfo.Reason)"
            Write-Log "Error invoking powershell thread. Reason: $($Powershell.InvocationStateInfo.Reason)"
            Throw "Error from the powershell first Instance"
        }
    }
#endregion ErrChk

    $Job = "" | Select-Object ID, Handle, Thread, Object, ParamsDict
    $Job.ID = $i
    if($i -eq 1 -and (-not $Force)){ #The first job is already invoked in the Process block if $force is not used
        $Job.Handle = $Handle
        $job.Thread = $Powershell
    }
    $Job.ParamsDict = $paramused
    $Job.Object = $Inputobject.ToString()
    $Jobs.Add($Job.ID,$Job)

    If(-not $Quiet){
    Write-Progress -id 1 -Activity "Creating Job Object" -Status "Created $i Objects"
    }

    }
    catch{
        Write-Error "Caught Error : $_"
        Write-Log "Caught Error : $_"
        Write-Host "Please verify if mandatory parameters are mentioned, if Command accepts ValueFromPipeline, do not explicity mention it as a Parameter." -ForegroundColor Yellow
        Write-Verbose "Cleaning up; Error in Process block"
        if($Powershell){
            $Powershell.dispose()
        }
        $runspacepool.Close()
        $runspacepool.dispose()

        if(Get-Command -CommandType function |?{$_.name -eq "$ProxyCommand"}){
            Remove-Item function:\$ProxyCommand -Confirm:$false    
            Write-Verbose "Deleted the Proxy function $ProxyCommand"
        }
        if(Get-Event |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'}){ 
            Get-Event |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'} | Remove-Event
        }
        if(Get-EventSubscriber |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'}){
            Get-EventSubscriber |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'} | Unregister-Event
        }
        
        Write-Verbose "Disposed the Powershell runspace pool objects"
        
        Break
        
    }
    finally{
    #nothing to do here, this block executes for all objects
     
    }

}#End Process Block


End{ #Invoke and collect the jobs as per the BatchSize, add pause if mentioned
    If(-Not $Quiet){
    Write-Progress -id 1 -Activity "Creating Job Object" -Completed
    }
    $Timer.Stop()
    $TotalObjects = $Jobs.Count

    Write-Verbose "Time took to create the Jobs : $($Timer.Elapsed.ToString())"
    Write-Log "Time took to create the Jobs : $($Timer.Elapsed.ToString())"
    Write-Log "Total Jobs created $TotalObjects"
    
    $Timer.Reset()

    [INT]$JobNum = 
    $Timer.Start()
    $SubTimer = [system.diagnostics.stopwatch]::StartNew() #SubTimer used to calculate time for batches

    
try{

#region begin ProcessJobs
#Start to invoke the jobs, first invoke Batchsize * 2 Jobs and then collect Batchsize jobs. This way we always Queue one batch in running state while collecting the completed jobs.

    #foreach($Job in $Jobs.Values.GetEnumerator() | Sort ID){
    for($JobNum = 1 ; $JobNum -le $TotalObjects ; $JobNum++){ #For loop is faster
#        $JobNum++
        $Job = $Jobs.Item($JobNum)
        if($Job.ID -eq 1 -and (-not $Force)){ Continue} #Skip the first job as it already invoked if Force switch is not specified
        $Powershell = [powershell]::Create()
    
        if($Commandtype -eq 'ExternalScript'){
            [VOID]$Powershell.AddScript($Code)
        }else{
            [VOID]$Powershell.AddCommand($Global:Command)
        }
        foreach($item in $Job.ParamsDict.GetEnumerator()){

            $Powershell.AddParameter($item.Key,$item.value) | Out-Null
        }

        $Powershell.RunspacePool = $RunspacePool
        $Job.Thread = $Powershell
        $Job.Handle = $Job.Thread.BeginInvoke()
        
        If(-Not $Quiet){
        Write-Progress -id 2 -Activity "Running Jobs" `
            -Status "$JobNum / $TotalObjects Jobs Invoked." `
            -PercentComplete $(($JobNum / $TotalObjects)*100)

        }    #-Status "$JobNum / $TotalObjects Jobs Invoked. $(@($Jobs |?{$_.Thread.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Running}).Count) Jobs are in Running State..." `

        if($JobNum -ge $JobCounter){ #Batch size is reached, lets start collecting the completed jobs

            $SubTimer.Stop()
            Write-Log "Time spent on Invoking Batch NO: $($Jobnum / $BatchSize) - $($SubTimer.Elapsed.ToString())"
            $SubTimer.Reset();$SubTimer.Start()

            CollectJobs $Jobs $BatchSize ([REF]$Script:jobsCollected)
            Write-Log "Jobs Collected so far: $script:jobsCollected"


            $SubTimer.Stop()
            Write-Log "Time spent on Collecting Batch NO: $($Jobnum / $BatchSize) - $($SubTimer.Elapsed.ToString())"
            $SubTimer.Reset()


            if($PauseInMsec){ #Pause before continuing to avoid throttling issues.
                Write-Verbose "Sleeping for $PauseInMsec Msecs"
                Start-Sleep -Milliseconds $PauseInMsec
            }

            $SubTimer.Start()

            $JobCounter += $BatchSize
            Write-Verbose "JobCounter : $JobCounter"
        
        }
    
    }

    $JobNum = $JobNum - 1 #For loops increments the num by 1 on exit, so rest it back
    Write-Verbose "All Jobs are invoked at this time, collecting all of them"
    Write-Log "Invoked all Jobs, Collecting the last jobs that are running"

    While (@($Jobs.Values.Handle | Where-Object {$_ -ne $Null}).count -gt 0)  {
        $BatchSize = @($Jobs.Values.Handle | Where-Object {$_ -ne $Null}).count * 2 #We want to collect all the Jobs, so just double the BatchSize
        Write-Log "Using BatchSize: $BatchSize"
        CollectJobs $Jobs $BatchSize ([ref]$Script:jobsCollected)
    }

    $Timer.Stop()
    Write-Verbose "Time took to Invoke and Complete the Jobs : $($Timer.Elapsed.ToString())"
    Write-Log "Jobs Collected: $script:jobsCollected"
    Write-Log "Time took to Invoke and Complete the Jobs : $($Timer.Elapsed.ToString())"
    
    $Timer.Reset()

#endregion ProcessJobs

#>
}Catch{
        Write-Error "Error executing jobs : $_ "

}



Finally{

    #Cleanup
    If(-Not $Quiet){
    Write-Progress -id 22  -ParentId 2 -Activity "Collecting Jobs results that are completed, $BatchSize at a time" -Completed
	Write-Progress -id 2 -Activity "Running Jobs" -Completed
    }
    if($(($Jobs.Values.GetEnumerator() |?{$_.Handle -ne $NULL} | Measure-Object ).count) -gt 0){
        
        Write-Host "Terminating the $(($Jobs.Values.Handle |?{$_ -ne $NULL} | Measure-Object ).count) Job(s) that are still running." -ForegroundColor Red
        Foreach($Job in ($Jobs.Values.GetEnumerator() |?{$_.Handle -ne $NULL})){
        
            $Job.Thread.stop()
		    $Job.Thread.Dispose()
		    $Job.Thread = $Null
		    $Job.Handle = $Null
    
        }
    }

    if($Powershell){
        $Powershell.dispose()
    }
    $runspacepool.Close()
    $runspacepool.dispose()
    $Jobs.Clear()
    $Jobs = $NULL
    if(Get-Command -CommandType function |?{$_.name -eq "$ProxyCommand"}){
        Remove-Item function:\$ProxyCommand -Confirm:$false    
        Write-Verbose "Deleted the Proxy function $ProxyCommand"
    }
    if(Get-Event |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'}){ 
        Get-Event |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'} | Remove-Event
    }
    if(Get-EventSubscriber |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'}){
        Get-EventSubscriber |?{$_.SourceIdentifier -eq 'PSInvocationStateChanged'} | Unregister-Event
    }
    
    [gc]::Collect()

    Write-Verbose "Triggered GC, Script execution has completed"
}
  
} #End End Block

}#End Function