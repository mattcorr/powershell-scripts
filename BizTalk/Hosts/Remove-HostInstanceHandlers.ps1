<# 
.SYNOPSIS 
    Will remove all Hosts, Host Instances and Handlers for the adapters as defined in a provided config.xml file
    
.DESCRIPTION 
    Given a valid XML file with defined hosts, host instances and adapters, this script will remove them all
    This will save users a lot of time as they wont have to manually remove them all via the BizTalk 
    Admin console.

.PARAMETER ConfigFile 
   the name of the config file to use. Contains info about hosts, host instancea and handlers to create  
     
.EXAMPLE 
    .\Remove-HostInstanceHandlers.ps1 -ConfigFile .\ClientConfig.xml
    Will create all hosts, host instances and handlers as defined in ClientConfig.xml
#> 
param (
    [string]$configurationFile = "$pwd\HostConfig.xml" 
)

if ((Test-Path -Path .\BiztalkScripts.psm1 -PathType Leaf) -eq $false)
{
    Write-Error "Unable to fine the BiztalkScripts.psm1 file. Check the folder and try again." -ErrorAction Stop
}
# import in the biztalk scripts needed (ignore the warnings)
Import-Module -Force -Name .\BizTalkScripts.psm1 -DisableNameChecking

# load the xml file
[xml] $xmldata = get-content $configurationFile

# go through all the host names
foreach ($hn in $xmlData.AdapterMappingConfiguration.HostNames.Host)
{
    if ($hn.tracking -eq "false")
    {
        # for each of the adapters
        foreach ($ad in $xmlData.AdapterMappingConfiguration.Adapters.Adapter)
        {
            if ($hn.isolated -eq "false")
            {		
                if ($ad.send -eq "true")
                {
                    Delete-Bts-Send-Handler $ad.name $hn.name
                }
                if ($ad.receive -eq "true")
                {
                    Delete-Bts-Receive-Handler $ad.name $hn.name
                }
            }
            else
            {
                if ($ad.isoReceive -eq "true")
                {
                    Delete-Bts-Receive-Handler $ad.name $hn.name
                }
            }
        }
    }
}