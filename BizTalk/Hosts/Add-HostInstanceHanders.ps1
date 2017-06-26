<# 
.SYNOPSIS 
    Will create all Hosts, Host Instances and Handlers for the adapters as defined in a provided config.xml file
    
.DESCRIPTION 
    Given a valid XML file with defined hosts, host instances and adapters, this script will create them all
    This will save users a lot of time as they wont have to manually create them all via the BizTalk 
    Admin console.
    There are switchs you can include to skip over parts of the scripts if desired.

.PARAMETER ConfigFile 
   the name of the config file to use. Contains info about hosts, host instancea and handlers to create  
.PARAMETER SkipHosts 
   If set, the Host creation section is skipped
.PARAMETER SkipHostInstances 
   If set, the Host Instance section creation is skipped
.PARAMETER SkipHostHandlers 
   If set, the Host Handler creation is skipped
     
.EXAMPLE 
    .\Add-HostInstanceHandlers.ps1 -ConfigFile .\ClientConfig.xml
    Will create all hosts, host instances and handlers as defined in ClientConfig.xml
    
    .\Add-HostInstanceHandlers.ps1 -ConfigFile .\ClientConfig.xml -SkipHosts -SkipHostInstances
    Will just create the handlers as defined in ClientConfig.xml
#> 

param (
    
   [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
   [string] $ConfigFile = "$pwd\HostConfig.xml",
   
   [switch] $SkipHosts,
   
   [switch] $SkipHostInstances,
   
   [switch] $SkipHostHandlers
    )

# ensure the Biztalk module is there
if ((Test-Path -Path .\BiztalkScripts.psm1 -PathType Leaf) -eq $false)
{
    Write-Error "Unable to fine the BiztalkScripts.psm1 file. Check the folder and try again." -ErrorAction Stop
}
# import in the biztalk scripts needed (ignore the warnings)
Import-Module -Force -Name .\BizTalkScripts.psm1 -DisableNameChecking

Write-Host "Loading configuration file: $configurationFile" -f DarkCyan
# load the xml file
[xml]$xmldata = Get-Content $configurationFile

Write-Host "Creating Hosts and Host Instances..." -f DarkCyan
# go through all the host names
#hostType = 1 = inProcess, 2 = Isolated
foreach ($hn in $xmlData.AdapterMappingConfiguration.HostNames.Host)
{
    if ($SkipHosts -eq $false)
    {
        Write-Host "Creating new Host: " -f DarkCyan -NoNewline; Write-Host $hn.name -f Cyan
        [int]$hostType = 1
        if ($hn.isolated -eq "true")
        {
            [int]$hostType = 2
        }
	    [bool]$trustBool = [Convert]::ToBoolean($hn.trusted)
	    [bool]$trackBool = [Convert]::ToBoolean($hn.tracking)
	    [bool]$is32BitBool = [Convert]::ToBoolean($hn.Is32Bit)
        $output = Create-Bts-Host $hn.name $hostType "$($hn.group)" $trustBool $trackBool $is32BitBool 
	}
    else
    {
        Write-Host "Skipping creation of Hosts." -f Yellow
    }

    if ($SkipHostInstances -eq $false)
    {
	    foreach ($hi in $hn.Instance)
	    {
            Write-Host "Creating instance of $($hn.name) on server $($hi.server)" -f DarkCyan
		    $output = Create-Bts-Instance $hn.name "$($hi.account)" $hi.password $hi.server
	    }
    }
    else
    {
        Write-Host "Skipping creation of Host Instances." -f Yellow
    }
    
	# for each of the adapters
    if (($SkipHostHandlers -eq $false) -and ($hn.tracking -eq "false"))
    {
        Write-Host "Creating Handlers for host $($hn.name)..." -f DarkCyan
	    foreach ($ad in $xmlData.AdapterMappingConfiguration.Adapters.Adapter)
	    {
		    if ($hn.isolated -eq "false")
		    {		
			    if ($ad.send -eq "true")
			    {
				    $output = Create-Bts-SendHandler $ad.name $hn.name
			    }
			    if ($ad.receive -eq "true")
			    {
				    $output = Create-Bts-ReceiveHandler $ad.name $hn.name
			    }
		    }
		    else
		    {
			    if ($ad.isoReceive -eq "true")
			    {
				    $output = Create-Bts-ReceiveHandler $ad.name $hn.name
			    }
		    }
	    }
    }
    else
    {
        Write-Host "Skipping creation of Hosts Handlers." -f Yellow
    }
}
Write-Host "All done!" -f DarkCyan
