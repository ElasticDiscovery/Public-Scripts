## 
# Setup: Log into a machine as the Relativity Service Account. There must be one machine to one environment; do NOT run this script for multiple environments on the same machine. 
#   An unseal event will cause ALL Reincarnation tasks to run, even other environments.
# Required information: Secretstore key, Secretstore hostname, Service account credentials, SQL Server instance
# Usage: Run this script in ISE, then use the functions Set-AutoSecretStoreUnseal and Enable-RelativityReincarnation.
# What it does: The two functions will create a task in the task scheduler to execute the scripts at the top of the file. The scripts are put to file on function execution.
# UnsealScript will check if the Secret Store service is down, or the store is sealed, and if so start it then trigger a forced Reincarnation.
# ReincarnateScript in it's default mode will connect to the designated SQL server, get the server manifest, then try to start the appropriate actions for each. If they are already started nothing will happen.
# If ReincarnateScript is ran with the -Force parameter, the above will happen but it will stop the services first then start them.
# Both scripts will log to text files that are generally in c:\powershell unless the ScriptPath parameter is overrode.
#
# Considerations: If your SQL server is down, this script will not help. There is currently no logic to detect if a service is "Running" but not responding. 
## 

$UnsealScript =
{
    param($SecretStoreServer)
    $Key = (Import-Clixml -Path "${env:\userprofile}\secretstore.xml").GetNetworkCredential().password 
    If (!($Key)) {
        "Key could not be loaded for secret store $($SecretStoreServer), most likely wrong user account."
        return
    }
    $Result = "Set"
    $ScriptText = 
@'
Push-Location "C:\Program Files\Relativity Secret Store\Client"
$Status = .\secretstore.exe seal-status | select -Skip 4 -First 1

switch ($Status)
{
    "Unable to connect to the remote server" {
        Start-Service "Relativity Secret Store"
        Start-Sleep -Seconds 8  
    }
    "Your store is currently unsealed and available to store and retrieve secrets." {
        
        return "Secret store is unsealed."
    }
}

$Status = .\secretstore.exe seal-status | select -Skip 4 -First 1
if($status -eq "Your store is sealed and requires the unseal key.")
{

    $Status = (.\secretstore.exe unseal _KEY_) | select -Skip 6 -First 1

    if($Status -ne "Success!")
    {
        return "Failed to unseal $($env:Computername), something is wrong with the key or database. f $($Status)"
    }
   
    return "Unsealed secret store."

}
'@

    $ScriptText = $ScriptText.Replace("_KEY_", $key)
    $ScriptBlock = [scriptblock]::Create($ScriptText)
    $Result = Invoke-Command -ComputerName $SecretStoreServer -ScriptBlock $ScriptBlock
    # Attempt to cause a forced Rel restart if the seal is unsealed. Generally an environment needs to be restarted after an event that causes a secret store to restart.
    if ($Result -eq "Unsealed secret store.") {
        # Reincarnate will created a scheduled task with this suffix.  
        if ($ReincarnateTask = get-scheduledtask "*_Reincarnate" -taskpath "\") {
            $SqlServer = $ReincarnateTask.TaskName.Replace("_Reincarnate", "")
            $ReincanateScriptDir = Split-Path $ReincarnateTask.Actions.Arguments
            if (Test-Path "$($ReincanateScriptDir)\RelativityReincarnation.ps1") {
                & "$($ReincanateScriptDir)\RelativityReincarnation.ps1" -SqlServer $Sqlserver -Force
                $Result = "Unsealed and force reincarnated the Relativity environment."
            }
            else {
                $Result = "Unsealed, task for reincarnate was present but script was not at $($ReincanateScriptDir)."
            }
        }
    }
    if ($UnsealTask = Get-ScheduledTask "*_Unseal" -TaskPath "\") {
        $UnsealScriptDir = Split-Path $UnsealTask.Actions.Arguments
        "$((Get-Date -Format 'yyyyMMddhhmmss')),$($SecretStoreServer),$($Result)" | Out-File "$($UnsealScriptDir)\AutoSecretStore.log" -Append
    }
}

$ReincarnateScript =
@'
param($SqlServer,[switch]$Force = $False)
# This script will call SQL to find all Rel servers in an environment, and attempt to start their services. If the Force flag is specified, they will be restarted.
# This function is lifted from Mr. Robbins to not have a dependency on sqlserver or sqlps modules. 
function Invoke-MrSqlDataReader {

<#
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>
 
    [CmdletBinding()]
    param (        
        [Parameter(Mandatory)]
        [string]$ServerInstance,
 
        [Parameter(Mandatory)]
        [string]$Database,
        
        [Parameter(Mandatory,
                   ValueFromPipeline)]
        [string]$Query,
        
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    BEGIN {
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
 
        if (-not($PSBoundParameters.Credential)) {
            $connectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;"
        }
        else {
            $connectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=False;"
            $userid= $Credential.UserName -replace '^.*\\|@.*$'
            ($password = $credential.Password).MakeReadOnly()
            $sqlCred = New-Object -TypeName System.Data.SqlClient.SqlCredential($userid, $password)
            $connection.Credential = $sqlCred
        }
 
        $connection.ConnectionString = $connectionString
        $ErrorActionPreference = 'Stop'
        
        try {
            $connection.Open()
            Write-Verbose -Message "Connection to the $($connection.Database) database on $($connection.DataSource) has been successfully opened."
        }
        catch {
            Write-Error -Message "An error has occurred. Error details: $($_.Exception.Message)"
        }
        
        $ErrorActionPreference = 'Continue'
        $command = $connection.CreateCommand()
    }
 
    PROCESS {
        $command.CommandText = $Query
        $ErrorActionPreference = 'Stop'
 
        try {
            $result = $command.ExecuteReader()
        }
        catch {
            Write-Error -Message "An error has occured. Error Details: $($_.Exception.Message)"
        }
 
        $ErrorActionPreference = 'Continue'
 
        if ($result) {
            $dataTable = New-Object -TypeName System.Data.DataTable
            $dataTable.Load($result)
            $dataTable
        }
    }
 
    END {
        $connection.Close()
    }
 
}
$GetServerQuery= 
@"
  select Name,Type from eddsdbo.extendedresourceserver
  where type in ('Agent','Services','Worker','Worker','Worker Manager Server')
  order by Type,Name
"@
try{

    $ReincarnateTask = Get-ScheduledTask "*_Reincarnate" -TaskPath "\"
    $ReincanateScriptDir = Split-Path $ReincarnateTask.Actions.Arguments

    if($Force){
     "Force restart started at $((Get-Date -Format 'yyyyMMddhhmmss'))" |out-file "$($ReincanateScriptDir)\Reincarnationlog.txt" -Append
    }
    else{
     "Reincarnation non-force started at  $((Get-Date -Format 'yyyyMMddhhmmss'))" |out-file "$($ReincanateScriptDir)\Reincarnationlog.txt" -append
    }

    $Servers = Invoke-MrSqlDataReader -ServerInstance $SqlServer -Database "EDDS" -query $GetServerQuery

    if(!$servers){
    "Unable to retrieve servers from database $((Get-Date -Format 'yyyyMMddhhmmss'))" |out-file "$($ReincanateScriptDir)\Reincarnationlog.txt" -Append
    }

    if($Agents = ($Servers |  ? Type -eq 'Agent').Name){
        $AgentJob = Invoke-Command -ComputerName $Agents -ScriptBlock{
            if($using:Force){
            stop-Service -Name 'kCura EDDS Agent Manager','kCura Service Host Manager'
            Start-Sleep -Seconds 5 
            }
            Start-Service -Name 'kCura EDDS Agent Manager','kCura Service Host Manager'
        } -AsJob -JobName 'Agent'
    }

    IF($Webs = ($Servers |  ? Type -eq 'Services').Name){
        $WebJob = Invoke-Command -ComputerName $Webs -ScriptBlock{
        Import-Module WebAdministration
            if($using:Force){    
                stop-Service -name 'kCura EDDS Web Processing Manager','kCura Service Host Manager' 
                IISReset /stop
                Start-Sleep -Seconds 5 
            
            }
            Start-Service -name 'kCura EDDS Web Processing Manager','kCura Service Host Manager' 
            IISReset /start
        
        } -AsJob -JobName 'Web'
}   
    if($WorkerManagers = ($Servers |  ? Type -eq 'Worker Manager Server').Name)
    {
        $WorkerManagerJob = Invoke-Command -ComputerName $WorkerManagers -ScriptBlock{
            if($using:Force){    
                stop-Service -name 'Invariant Queue Manager' 
                Start-Sleep -Seconds 5 
            }
            Start-Service -name 'Invariant Queue Manager' 
        } -AsJob -JobName 'InvariantQM'
    }

        If($Workers = ($Servers |  ? Type -eq 'Worker').Name){ 
        $WorkerJob = Invoke-Command -ComputerName $Workers -ScriptBlock{
           Start-ScheduledTask -TaskName 'Relativity Processing Launcher' -TaskPath 'Invariant\'
        } -AsJob -JobName 'InvariantWorker'
    }
}
catch{
    "Exception: $($PSItem.Exception.Message) ,$((Get-Date -Format 'yyyyMMddhhmmss'))" | Out-File "$($ReincanateScriptDir)\Reincarnationlog.txt" -Append
}
    # It's possible that these jobs don't exist, so we want to surpress these errors. 
    Wait-Job -Job $AgentJob -ErrorAction SilentlyContinue
    Wait-Job -Job $WebJob -ErrorAction SilentlyContinue
    Wait-Job -Job $WorkerManagerJob -ErrorAction SilentlyContinue
    Wait-Job -Job $WorkerJob -ErrorAction SilentlyContinue

if($Force){
 "Force restart finished at $((Get-Date -Format 'yyyyMMddhhmmss'))" | Out-File "$($ReincanateScriptDir)\Reincarnationlog.txt" -Append
}
else{
 "Reincarnation non-force finished at  $((Get-Date -Format 'yyyyMMddhhmmss'))" | Out-File "$($ReincanateScriptDir)\Reincarnationlog.txt" -append
}


'@


function Set-AutoSecretStoreUnseal {
    param(
        [Parameter(Mandatory = $true)]$SecretStoreServers,
        $ScriptPath = "C:\powershell\UnsealSecretStore.ps1")

    $Key = Get-Credential -UserName "secretstore" -Message "Enter the secret store key as the password."
    $KeyPath = "$(${env:\userprofile})\secretstore.txt"
    $ScriptDir = [System.IO.DirectoryInfo]$ScriptPath
    
    $ServiceAccount = Get-Credential -Message "Enter Service Account creds"

    if (test-path $keypath) {
        Read-Host "There is a key already present. Press enter to continue and override or ctrl+c to stop."
    }

    $Key | Export-Clixml  -Path "${env:\userprofile}\secretstore.xml"

    If (!(Test-path $ScriptDir.Parent.FullName)) {
        New-Item -ItemType Directory -path $ScriptDir.Parent.FullName
    }

    $UnsealScript | Out-File $ScriptPath 

    foreach ($Server in $SecretStoreServers) {

        $TaskName = "$($Server)_Unseal"
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $TaskName 

        }
        $Trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -day 5000)  -At (Get-Date) -Once
        $Action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "$($ScriptPath) -SecretStoreServer $($Server)" 
        Register-ScheduledTask -TaskName $TaskName -Trigger $Trigger -Action $action -User $ServiceAccount.UserName -Password $ServiceAccount.GetNetworkCredential().Password
    }

}

function Enable-RelativityReincarnation {
    param(
        [Parameter(Mandatory = $true)]$SQLServer,
        $ScriptPath = "C:\powershell\RelativityReincarnation.ps1")
    
    $ServiceAccount = Get-Credential -Message "Enter Service Account creds"
    $TaskName = "$($SqlServer)_Reincarnate"
    $ScriptDir = [System.IO.DirectoryInfo]$ScriptPath
    If (!(Test-path $ScriptDir.Parent.FullName)) {
        New-Item -ItemType Directory -path $ScriptDir.Parent.FullName
    }
    $ReincarnateScript | Out-File $ScriptPath 
    
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName 
    }
     
    $Trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 5000)  -At (Get-Date) -Once 
    $Action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "$($ScriptPath) -SQLServer $($SQLServer)" 
    Register-ScheduledTask -TaskName $TaskName -Trigger $Trigger -Action $action -User $ServiceAccount.UserName -Password $ServiceAccount.GetNetworkCredential().Password

}