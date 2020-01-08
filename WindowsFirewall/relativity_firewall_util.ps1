#requires -version 3
<#
.SYNOPSIS
    Helper utility to create windows firewall rules for various relativity components.
.PARAMETER targetRole
    The desired role for which to apply firewall rules.
.NOTES
    Version:        1.1
    Author:         Chris Eastwood
    Creation Date:  2019-12-12
    Purpose/Change: Initial script development
#>
param (
    # Target Relativity Role
    [Parameter()]
    [string]
    $targetRole,
    [Parameter()]
    [switch]
    $overwrite
)

$ErrorActionPreference = "Stop"
$InfoColor = [ConsoleColor]::Cyan
$WarnColor = [ConsoleColor]::Yellow
$DangerColor = [ConsoleColor]::Red
$DefaultColor = [ConsoleColor]::White
$SuccessColor = [ConsoleColor]::Green
$RolesJsonPath = 'relativity_firewall_util_roles.json'

function Log() 
{
    param (
        [Parameter()]
        [string]
        $message,
        [Parameter()]
        [ConsoleColor]
        $color = $DefaultColor
    )

    Write-Host $message -ForegroundColor $color -NoNewline
}

function LogLine() 
{
    param (
        [Parameter()]
        [string]
        $message,
        [Parameter()]
        [ConsoleColor]
        $color = $DefaultColor
    )

    Write-Host $message -ForegroundColor $color
}

LogLine "----------" $InfoColor
LogLine "Windows Firewall Utility for Relativity - https://www.elasticdiscovery.com" $InfoColor
Log "For comments, questions, or troubleshooting, please send us an email at " $InfoColor
LogLine "relativity@elasticdiscovery.com" $DefaultColor
LogLine "----------" $InfoColor
LogLine

function doExit($exitCode = 0) {
    if ($exitCode -ne 0) 
    {
        Log "Exiting with code: " $WarnColor
        LogLine $exitCode $DangerColor
    }
    exit $exitCode
}

function LogAndExit($message, $color = $DangerColor, $exitCode = 0) {
    LogLine $message $color
    doExit $exitCode
}

function AskConfirm() {
    Log "Proceed? [Y] to continue: " $WarnColor
    $confirmation = Read-Host
    if ($confirmation -eq 'y') {
        return $true
    }    
    return $false
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    LogAndExit "You are not a member of the Administrators group." $DangerColor 1
}

if (!(Test-Path $RolesJsonPath))
{
    LogAndExit "The source file for roles '$RolesJsonPath' was not found."
}

$relativityRoles = Get-Content $RolesJsonPath | ConvertFrom-Json

function LogHelpAndExit($exitCode) {
    $slimRoles = $relativityRoles | Select-Object -Property role, key | Format-Table | Out-String
    LogLine "-overwrite`tOverwrites any firewall rule with a matching name"
    LogLine "-targetRole <role>"
    LogAndExit $slimRoles White $exitCode
}

if ([string]::IsNullOrWhiteSpace($targetRole))
{
    LogLine "You must provide a valid role." $DangerColor
    LogHelpAndExit 3
}

$relativityRole = $relativityRoles | Where-Object {$_.key -eq $targetRole}
if (-not $relativityRole)
{
    LogLine "You have provided an invalid target role." $DangerColor
    LogHelpAndExit
}

Log "Target Role Key: "
LogLine $targetRole $InfoColor

$confirmed = AskConfirm
if ($confirmed -eq $false)
{
    doExit(4)
}

LogLine "Creating firewall rules..."
$ruleNameTemplate = "kCura Relativity {0} (ALLOW {1} {2})"

function createFirewallRule($ruleName, $direction, $ports, $protocol)
{
    try {
        $rule = Get-NetFirewallRule -DisplayName "$ruleName" 2> $null;
    }
    catch {
        $rule = $null
    }
    
    if ($rule -and $overwrite) {
        LogLine "`tRemoving existing rule: $ruleName" $WarnColor
        Remove-NetFirewallRule -InputObject $rule
        $rule = $null
    }

    $cmd = "New-NetFirewallRule -DisplayName `"$ruleName`" -Protocol $protocol -LocalPort $ports -Direction $direction -Action Allow | Out-Null"
    if (-not $rule)
    {
        try {
            Invoke-Expression -Command $cmd 2>$null;
        }
        catch {
            LogLine "`tFAILED: '$ruleName'" $DangerColor
        }
        LogLine "`tCreated: '$ruleName'" $SuccessColor
    } else {
        LogLine "`tSkipped: '$ruleName' exists" $WarnColor
    }
}

function LogSkippedNoPorts($ruleName)
{
    LogLine "`tSkipped: '$ruleName' no ports to allow" $WarnColor
}

$ruleName = $ruleNameTemplate -f $relativityRole.role,"Inbound", "Tcp"
if (-not [string]::IsNullOrWhiteSpace($relativityRole.inboundTcp))
{
    $ports = $relativityRole.inboundTcp
    createFirewallRule $ruleName 'Inbound' $ports 'Tcp'
} else {
    LogSkippedNoPorts $ruleName
}

$ruleName = $ruleNameTemplate -f $relativityRole.role,"Outbound", "Tcp"
if (-not [string]::IsNullOrWhiteSpace($relativityRole.outboundTcp))
{
    
    $ports = $relativityRole.outboundTcp
    createFirewallRule $ruleName 'Outbound' $ports 'Tcp'
} else {
    LogSkippedNoPorts $ruleName
}

$ruleName = $ruleNameTemplate -f $relativityRole.role,"Inbound", "Udp"
if (-not [string]::IsNullOrWhiteSpace($relativityRole.inboundUdp))
{
    $ports = $relativityRole.inboundUdp
    createFirewallRule $ruleName 'Inbound' $ports 'Udp'
} else {
    LogSkippedNoPorts $ruleName
}

$ruleName = $ruleNameTemplate -f $relativityRole.role,"Outbound", "Udp"
if (-not [string]::IsNullOrWhiteSpace($relativityRole.outboundUdp))
{
    $ports = $relativityRole.outboundUdp
    createFirewallRule $ruleName 'Outbound' $ports 'Udp'
} else {
    LogSkippedNoPorts $ruleName
}

doExit 0





