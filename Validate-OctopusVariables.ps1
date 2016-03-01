<#
.SYNOPSIS
    Parses all Octopus Deploy variables for Applications and Library Sets and checks for any duplicates. (Variable name and Environment)
    Will also take into account Machines defined in scope as well. Other scoping properties (Roles and Steps) are NOT supported at this time.
 
.DESCRIPTION
    It is a good idea to ensure that there is only one definition of a variable for each environment in the Octopus variable configuration for either an Application
    or a Library Set. If there is more than one accidently defined, it is undetermined which one is actually used during the deployment itself. (Confirmed by Octopus developers)
    This script will highlight any duplicates detected.
    It also takes into account if Machines are defined in the scope as well. This means (for example) if a variable is defined twice with the same environment but
    DIFFERENT machines then it is not classed as a duplicate.

 
.PARAMETER ServerName
   The name of the Octopus Deploy Server.If not provided it will default to the current instance.
.PARAMETER ApiKey
   The Octopus Deploy API key for the user running this script. If not provided it will try to use $OctopusApiKey
 
.EXAMPLE
   .\Validate-OctopusVariables.ps1
    This command uses the default config parameters and will display any duplicates found in the Octopus variable database
#>
param (
    [string]$ServerName = "Put your Octopus Server here",
   
    [string]$ApiKey = $OctopusApiKey
 
    )
 

function Get-LibraryVariableSets($Server, $ApiKey)
{
    return Invoke-RestMethod -Uri "$Server/api/libraryvariablesets/all" -Headers @{"X-Octopus-ApiKey"=$ApiKey}
}

# ------------------------------------------------------------------------------------------------------

function Get-Projects($Server, $ApiKey)
{
    return Invoke-RestMethod -Uri "$Server/api/projects/all" -Headers @{"X-Octopus-ApiKey"=$ApiKey}
}

# ------------------------------------------------------------------------------------------------------

function Get-Projects($Server, $ApiKey)
{
    return Invoke-RestMethod -Uri "$Server/api/projects/all" -Headers @{"X-Octopus-ApiKey"=$ApiKey}
}

# ------------------------------------------------------------------------------------------------------
# MAIN
# ======================================================================================================
$ProgressPreference="SilentlyContinue"
 
Write-Host "Reading all Octopus Projects..." -f DarkCyan
$projects = Get-Projects -Server $ServerName -ApiKey $ApiKey
 
$allProjectVariables = @()
 
Write-Host "Checking Application Variables for duplicate values..."
foreach ($app in $projects)
{
    Write-Host "$($app.name): " -f Cyan -NoNewline
    # get the variable
    $appVars = Get-VariableSet $ServerName $ApiKey $app.VariableSetId
 
    $dupeValues = @()
    foreach ($env in $appVars.ScopeValues.Environments)
    {
        $dupesTemp = ($appVars.Variables | Where { $_.Scope.Environment -Contains $env.Id} | Group-Object -Property Name) | where Count -gt 1
        if ($dupesTemp)
        {
            foreach ($dupe in $dupesTemp)
            {
                # this additional filter will check to see if the scopes have different machines defined. (if so then the same Env is okay)
                if (($dupe.Group.Scope.Machine | select -Unique).count -lt 2)  # needs to be 0 or 1 to be a classed as a dupe
                {
                    $dupeValues += @{ Data = $dupe; EnvName = $env.name}
                }
            }
        }
    }
    if ($dupeValues)
    {
        Write-Host "$($dupeValues.Count) duplicate$(if ($dupeValues.Count -ne 1) {"s"}) found" -f Red
        foreach ($dupe in $dupeValues)
        {
            Write-Host " - Variable: $($dupe.Data.Name) with Environment: $($dupe.EnvName) (Count: $($dupe.Data.Count))." -f Yellow
        }
        Write-Host
    }
    else
    {
        Write-Host " OK" -f Green
    }
}
Write-Host
Write-Host
Write-Host "Checking Library Set Variables for duplicate values..."
 
$libVarSets = Get-LibraryVariableSets $ServerName $ApiKey
 
foreach ($libVarSet in $libVarSets)
{
    Write-Host "$($libVarSet.Name): " -NoNewline -f Cyan
 
    $libVar = Get-VariableSet $ServerName $ApiKey $libVarSet.VariableSetId
 
    $dupeValues = @()
    foreach ($env in $libVar.ScopeValues.Environments)
    {
        $dupesTemp = ($libVar.Variables | Where { $_.Scope.Environment -Contains $env.Id} | Group-Object -Property Name) | where Count -gt 1
       
        foreach ($dupe in $dupesTemp)
        {
            # this additional filter will check to see if the scopes have different machines defined. (if so then the same Env is okay)
            if (($dupe.Group.Scope.Machine | select -Unique).count -lt 2)  # needs to be 0 or 1 to be a classed as a dupe
            {
                $dupeValues += @{ Data = $dupe; EnvName = $env.name}
            }
        }
    }
    if ($dupeValues)
    {
        Write-Host "$($dupeValues.Count) duplicate$(if ($dupeValues.Count -ne 1) {"s"}) found" -f Red
        foreach ($dupe in $dupeValues)
        {
            Write-Host " - Variable: $($dupe.Data.Name) with Environment: $($dupe.EnvName) (Count: $($dupe.Data.Count))." -f Yellow
        }
        Write-Host
    }
    else
    {
        Write-Host " OK" -f Green
    }
}
Write-Host "All done." -f DarkCyan
