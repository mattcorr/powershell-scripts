function write-WarnMessage([string]$message) {
    Write-Host $(Get-Date) $message -ForegroundColor Yellow
}
function write-SucessMessage([string]$message) {
    Write-Host $(Get-Date) $message -ForegroundColor Green
}
function write-InfoMessage([string]$message) {
    Write-Host $(Get-Date) $message -ForegroundColor Blue -BackgroundColor White
}
function write-ErrorMessage ([string]$message) {
    Write-Host $(Get-Date) $message -ForegroundColor Red
}
# Gets the execution directory
function Get-ScriptDirectory 
{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

function SetAsDefaultHost-Bts-Send-Handler ( [string]$adapter, [string]$hostName )
{
    try
    {
        [System.Management.ManagementObject]$objHandler = get-wmiobject 'MSBTS_SendHandler2' -namespace 'root\MicrosoftBizTalkServer' -filter "HostName='$hostName' AND AdapterName='$adapter'"
        $objHandler["IsDefault"] = $true
		$objHandler.Put()
	
        write-SucessMessage "Set $hostName as Default Host for $adapter"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        if ($_.Exception.Message -eq "You cannot call a method on a null-valued expression.")
        {
            write-WarnMessage "$adapter send handler for $hostName does not exist"
        }
        else
        {
            write-Error "$adapter send handler for $hostName could not be deleted: $_.Exception.ToString()"
        }
    }
}

function Delete-Bts-Receive-Handler ( [string]$adapter, [string]$hostName )
{   
    try
    {
        [System.Management.ManagementObject]$objHandler = get-wmiobject 'MSBTS_ReceiveHandler' -namespace 'root\MicrosoftBizTalkServer' -filter "HostName='$hostName' AND AdapterName='$adapter'"
        $objHandler.Delete()
        write-SucessMessage "Deleted $adapter receive handler for $hostName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        if ($_.Exception.Message -eq "You cannot call a method on a null-valued expression.")
        {
            write-WarnMessage "$adapter receive handler for $hostName does not exist"
        }
        elseif ($_.Exception.Message.IndexOf("Cannot delete a receive handler that is used by") -ne -1)
        {
            write-WarnMessage "$adapter receive handler for $hostName is in use. Cannot delete."
        }
        else
        {
            write-Error "$adapter receive handler for $hostName could not be deleted: $_.Exception.ToString()"
        }
    }
}
function Delete-Bts-Send-Handler ( [string]$adapter, [string]$hostName )
{
    try
    {
        [System.Management.ManagementObject]$objHandler = get-wmiobject 'MSBTS_SendHandler2' -namespace 'root\MicrosoftBizTalkServer' -filter "HostName='$hostName' AND AdapterName='$adapter'"
        $objHandler.Delete()
        write-SucessMessage "Deleted $adapter send handler for $hostName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        if ($_.Exception.Message -eq "You cannot call a method on a null-valued expression.")
        {
            write-WarnMessage "$adapter send handler for $hostName does not exist"
        }
        elseif ($_.Exception.Message.IndexOf("Cannot delete a send handler that is used by") -ne -1)
        {
            write-WarnMessage "$adapter send handler for $hostName is in use. Cannot delete."
        }
        else
        {
            write-Error "$adapter send handler for $hostName could not be deleted: $_.Exception.ToString()"
        }
    }
}
function Delete-Bts-Instance( [string]$hostName, [string]$Server )
{
    try
    {
        # Unintall
        [System.Management.ManagementObject]$objHostInstance = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostInstance").CreateInstance()
        $name = "Microsoft BizTalk Server " + $hostName + " " + $Server
        $objHostInstance["Name"] = $name
        $objHostInstance.Uninstall()
        # Unmap
        [System.Management.ManagementObject]$objServerHost = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_ServerHost").CreateInstance()
        $objServerHost["HostName"] = $hostName
        $objServerHost["ServerName"] = $Server
        $objServerHost.Unmap()
        write-SucessMessage "Deleted host instance for $hostName on $Server"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host instance on server $Server could not be deleted: $_.Exception.ToString()"
    }
}
function Stop-Bts-HostInstance ( [string] $HostName, [string] $Server )
{
    try
    {
        $filter = "HostName = '" + $HostName + "' and RunningServer = '" + $Server + "'"
        $HostInstance = get-wmiobject "MSBTS_HostInstance" -namespace "root\MicrosoftBizTalkServer" -filter $filter
        $HostInstanceState = $HostInstance.GetState().State
        write-InfoMessage "Current state of $HostName instance on server $Server : $HostInstanceState (1=Stopped, 2=Start pending, 3=Stop pending, 4=Running, 8=Unknown)"
        if ($HostInstanceState -eq 4) 
        {
            $HostInstance.Stop() 
            $HostInstanceState = $HostInstance.GetState().State
            write-SucessMessage "New state of $HostName instance on server $Server : $HostInstanceState (1=Stopped, 2=Start pending, 3=Stop pending, 4=Running, 8=Unknown)"
        }
        else
        {
            $HostInstanceState = $HostInstance.GetState().State
            write-WarnMessage "Failed to stop host instance $HostName on server $Server because host instance state $HostInstanceState was not the expected value of 4 (running)"
        }
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host instance could not be stopped on $Server : $_.Exception.ToString()"
    }
}
function Delete-Bts-Host ( [string]$hostName )
{
    # TODO: This only works intermittently
    try
    {
        [System.Management.ManagementObject]$objHostSetting = get-wmiobject 'MSBTS_HostSetting' -namespace 'root\MicrosoftBizTalkServer' -filter "HostName='$hostName'"
        $objHostSetting.Delete()
        write-SucessMessage "Deleted host $hostName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host could not be deleted: $_.Exception.ToString()"
    }
}


# function to create BizTalk host
# For HostType 1 = inProcess, 2 = Isolated
function Create-Bts-Host(
    [string]$hostName, 
    [int]$hostType, 
    [string]$ntGroupName, 
    [bool]$authTrusted, 
    [bool]$tracking, 
    [bool]$32BitOnly)
{
    try
    {
        [System.Management.ManagementObject]$objHostSetting = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostSetting").CreateInstance()
        $objHostSetting["Name"] = $hostName
        $objHostSetting["HostType"] = $hostType
        $objHostSetting["NTGroupName"] = $NTGroupName
        $objHostSetting["AuthTrusted"] = $authTrusted
        $objHostSetting["IsHost32BitOnly"] = $32BitOnly 
        $objHostSetting["HostTracking"] = $tracking
        $putOptions = new-Object System.Management.PutOptions
        $putOptions.Type = [System.Management.PutType]::CreateOnly;
        [Type[]] $targetTypes = New-Object System.Type[] 1
        $targetTypes[0] = $putOptions.GetType()
        $sysMgmtAssemblyName = "System.Management"
        $sysMgmtAssembly = [System.Reflection.Assembly]::LoadWithPartialName($sysMgmtAssemblyName)
        $objHostSettingType = $sysMgmtAssembly.GetType("System.Management.ManagementObject")
        [Reflection.MethodInfo] $methodInfo = $objHostSettingType.GetMethod("Put", $targetTypes)
        $methodInfo.Invoke($objHostSetting, $putOptions)
        write-SucessMessage "Host $hostName created"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host could not be created: $_.Exception.ToString()"
    }
}

# For HostType 1 = inProcess, 2 = Isolated
function Update-Bts-Host ( 
    [string]$hostName, 
    [int]$hostType, 
    [string]$ntGroupName, 
    [bool]$authTrusted, 
    [bool]$tracking, 
    [bool]$32BitOnly)
{
    try
    {
        [System.Management.ManagementObject]$objHostSetting = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostSetting").CreateInstance()
        $objHostSetting["Name"] = $hostName
        $objHostSetting["HostType"] = $hostType
        $objHostSetting["NTGroupName"] = $ntGroupName
        $objHostSetting["AuthTrusted"] = $authTrusted
        $objHostSetting["IsHost32BitOnly"] = $32BitOnly 
        $objHostSetting["HostTracking"] = $tracking
        $putOptions = new-Object System.Management.PutOptions
        $putOptions.Type = [System.Management.PutType]::UpdateOnly; # This tells WMI it's an update.
        [Type[]] $targetTypes = New-Object System.Type[] 1
        $targetTypes[0] = $putOptions.GetType()
        $sysMgmtAssemblyName = "System.Management"
        $sysMgmtAssembly = [System.Reflection.Assembly]::LoadWithPartialName($sysMgmtAssemblyName)
        $objHostSettingType = $sysMgmtAssembly.GetType("System.Management.ManagementObject")
        [Reflection.MethodInfo] $methodInfo = $objHostSettingType.GetMethod("Put", $targetTypes)
        $methodInfo.Invoke($objHostSetting, $putOptions)
        write-SucessMessage "Host updated"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host could not be updated: $_.Exception.ToString()"
    }
}
# function to create BizTalk send adapter handler
function Create-Bts-SendHandler([string]$adapter, [string]$hostName)
{
    try
    {
        [System.Management.ManagementObject]$objSendHandler = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_SendHandler2").CreateInstance()
        $objSendHandler["AdapterName"] = $adapter
        $objSendHandler["HostName"] = $hostName
        $objSendHandler["IsDefault"] = $false
        $putOptions = new-Object System.Management.PutOptions
        $putOptions.Type = [System.Management.PutType]::CreateOnly;
        [Type[]] $targetTypes = New-Object System.Type[] 1
        $targetTypes[0] = $putOptions.GetType()
        $sysMgmtAssemblyName = "System.Management"
        $sysMgmtAssembly = [System.Reflection.Assembly]::LoadWithPartialName($sysMgmtAssemblyName)
        $objSendHandlerType = $sysMgmtAssembly.GetType("System.Management.ManagementObject")
        [Reflection.MethodInfo] $methodInfo = $objSendHandlerType.GetMethod("Put", $targetTypes)
        $methodInfo.Invoke($objSendHandler, $putOptions)
        write-SucessMessage "Send handler created for $adapter / $hostName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "Send handler for $adapter / $hostName could not be created: $_.Exception.ToString()"
    }
}
# function to create BizTalk receive adapter handler
function Create-Bts-ReceiveHandler([string]$adapter, [string]$hostName)
{
    try
    {
        [System.Management.ManagementObject]$objReceiveHandler = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_ReceiveHandler").CreateInstance()
        $objReceiveHandler["AdapterName"] = $adapter
        $objReceiveHandler["HostName"] = $hostName
        $putOptions = new-Object System.Management.PutOptions
        $putOptions.Type = [System.Management.PutType]::CreateOnly;
        [Type[]] $targetTypes = New-Object System.Type[] 1
        $targetTypes[0] = $putOptions.GetType()
        $sysMgmtAssemblyName = "System.Management"
        $sysMgmtAssembly = [System.Reflection.Assembly]::LoadWithPartialName($sysMgmtAssemblyName)
        $objReceiveHandlerType = $sysMgmtAssembly.GetType("System.Management.ManagementObject")
        [Reflection.MethodInfo] $methodInfo = $objReceiveHandlerType.GetMethod("Put", $targetTypes)
        $methodInfo.Invoke($objReceiveHandler, $putOptions)
        write-SucessMessage "Receive handler created for $adapter / $hostName"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "Receive handler for $adapter / $hostName could not be created: $_.Exception.ToString()"
    }
}
# function to create BizTalk host instance
function Create-Bts-Instance([string]$hostName, [string]$login, [string]$password, [string]$Server)
{
    try
    {
        [System.Management.ManagementObject]$objServerHost = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_ServerHost").CreateInstance()
        $objServerHost["HostName"] = $hostName
        $objServerHost["ServerName"] = $Server
        $objServerHost.Map()
        [System.Management.ManagementObject]$objHostInstance = ([WmiClass]"root/MicrosoftBizTalkServer:MSBTS_HostInstance").CreateInstance()
        $name = "Microsoft BizTalk Server " + $hostName + " " + $Server
        $objHostInstance["Name"] = $name
        $objHostInstance.Install($Login, $Password, $True)
        write-SucessMessage "Created host instance for $hostName on $Server"
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host instance on server $Server could not be created: $_.Exception.ToString()"
    }
}
function Start-Bts-HostInstance ( [string] $HostName, [string] $Server )
{
    try
    {
        $filter = "HostName = '" + $HostName + "' and RunningServer = '" + $Server + "'"
        $HostInstance = get-wmiobject "MSBTS_HostInstance" -namespace "root\MicrosoftBizTalkServer" -filter $filter
        $HostInstanceState = $HostInstance.GetState().State
        write-InfoMessage "Current state of $HostName instance on server $Server : $HostInstanceState (1=Stopped, 2=Start pending, 3=Stop pending, 4=Running, 8=Unknown)"
        if ($HostInstanceState -eq 1) 
        {
            $HostInstance.Start() 
            $HostInstanceState = $HostInstance.GetState().State
            write-SucessMessage "New state of $HostName instance on server $Server : $HostInstanceState (1=Stopped, 2=Start pending, 3=Stop pending, 4=Running, 8=Unknown)"
        }
        else
        {
            $HostInstanceState = $HostInstance.GetState().State
            write-WarnMessage "Failed to start host instance $HostName on server $Server because host instance state $HostInstanceState was not the expected value of 1 (stopped)"
        }
    }
    catch [System.Management.Automation.RuntimeException]
    {
        write-Error "$hostName host instance could not be started on $Server : $_.Exception.ToString()"
    }
}
function Install-BTSMsi ( [string]$bts_application, [string]$msi_package, [string]$install_env ) 
{
    write-InfoMessage "Installing $msi_package in $bts_application for $install_env"
    BTSTask ImportApp /ApplicationName:$bts_application /Package:$msi_package /Overwrite /Environment:$install_env
    if ($LASTEXITCODE -ne 0) {
        write-ErrorMessage "Failed to Import MSI $msi_package"
    } 
    else
    {
        write-SucessMessage "Installed $bts_application for $install_env"
    }
}
function Remove-BTSApplication  ( [string]$appServer, [string]$appDatabase, [string]$appName ) 
{
    write-InfoMessage "Uninstalling Application: $appName "
    BTSTask RemoveApp /Server:"$appServer" /ApplicationName:"$appName" /Database:"$appDatabase"  #| out-null
    if ($LASTEXITCODE -ne 0) {
        write-ErrorMessage "Failed to remove $appServer $appName"
    }
    else
    {
        write-SucessMessage "Removed $appName from $appServer"
    }
} 
# Accesses SSO and will require the build user account to belong to the SSO Admins group. 
# Also requires Microsoft.BizTalk.ExplorerOM.dll to be loaded.
function StartStop-BTSApplication ( [string]$appServer, [string]$appName, [string]$appCommand ) 
{ 
    if ( ($appName -eq '') -or ($appName -eq $null) )
    {
        throw 'you must supply the application name'
    }
    #write-InfoMessage " Finding Application: $appServer:$appName "
    $exp = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer
    $exp.ConnectionString = Get-BTSConnectionString($appServer) 
    $app = $exp.Applications[$appName] 
    if($app -eq $null) 
    { 
        if ($appCommand -eq "stop" )
        {
            write-WarnMessage "WARNING failed to stop $appName"
        }
        else
        {
            write-WarnMessage "FAILED to start $appName"
        }
    } 
    else 
    {
        switch -wildcard ( $app.Status.ToString() )
        {
            'Stopped' 
            {
                if ($appCommand -eq "start" ) {
                    write-InfoMessage "Starting Application: $appName "
                    $null = $app.Start([Microsoft.BizTalk.ExplorerOM.ApplicationStartOption]::StartAll) 
                    $null = $exp.SaveChanges() 
                    write-SucessMessage " Started Application: $appName "
                } else {
                    write-InfoMessage "Application Already Stopped: $appName "
                }
            } 
            '*Started' 
            { 
                # includes Started and PartiallyStarted
                if ($appCommand -eq "stop" ) {
                    write-InfoMessage "Stopping Application: $appName "
                    $null = $app.Stop([Microsoft.BizTalk.ExplorerOM.ApplicationStopOption]::StopAll) 
                    $null = $exp.SaveChanges() 
                    write-SucessMessage " Stopped Application: $appName "
                } else {
                    write-InfoMessage "Application Already Started : $appName "
                }
            }
            'NotApplicable' 
            {
                write-InfoMessage "Application doesn't require $appCommand"
            } 
            default
            {
                $msg = "Unkown STATUS: " + $app.Status
                write-ErrorMessage $msg
            }
        }
    }
}

function Get-BTSConnectionString ( [string] $server )
{
    $group = Get-WmiObject MSBTS_GroupSetting -n root\MicrosoftBizTalkServer -computername $server
    $grpdb = $group.MgmtDBName
    $grpsvr = $group.MgmtDBServerName
    [System.String]::Concat("server=", $grpsvr, ";database=", $grpdb, ";Integrated Security=SSPI")
    write-InfoMessage " Server: $grpsvr - Database  $grpdb"
} 

