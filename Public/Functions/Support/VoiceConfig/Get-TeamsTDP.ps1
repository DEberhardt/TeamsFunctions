﻿# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:		David Eberhardt
# Updated:  01-JAN-2021
# Status:   Live




function Get-TeamsTDP {
  <#
  .SYNOPSIS
    Lists all Tenant Dial Plans by Name
  .DESCRIPTION
    To quickly find Tenant Dial Plans to assign, an Alias-Function to Get-CsTenantDialPlan
  .PARAMETER Identity
    If provided, acts as an Alias to Get-CsTenantDialPlan, listing one Dial Plan
    If not provided, lists Identities of all Tenant Dial Plans (except "Global")
  .EXAMPLE
    Get-TeamsTDP
    Lists Identities (Names) of all Tenant Dial Plans (except "Global")
  .EXAMPLE
    Get-TeamsTDP -Identity DP-HUN
    Lists Tenant Dial Plan DP-HUN as Get-CsTenantDialPlan does (provided it exists).
  .NOTES
    Without parameters, it executes the following string:
    Get-CsTenantDialPlan | Where-Object Identity -NE "Global" | Select-Object Identity -ExpandProperty Identity
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Get-TeamsTDP
  .LINK
    Get-TeamsVNR
  .LINK
    Get-TeamsOVP
  .LINK
    Get-TeamsOPU
  .LINK
    Get-TeamsOVR
  .LINK
    Get-TeamsMGW
  #>

  [CmdletBinding()]
  param (
    [Parameter(Position = 0, ValueFromPipeline, HelpMessage = "Name of the Tenant Dial Plan")]
    [string]$Identity
  )

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"

    # Asserting SkypeOnline Connection
    if (-not (Assert-SkypeOnlineConnection)) { break }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    if ($PSBoundParameters.ContainsKey('Identity')) {
      Write-Verbose -Message "Finding Tenant Dial Plans with Identity '$Identity'"
      $Plans = Get-CsTenantDialPlan -WarningAction SilentlyContinue
      $Filtered = $Plans | Where-Object Identity -Like "*$Identity*"
      if ( $Filtered.Count -gt 2) {
        $Filtered | Select-Object Identity
      }
      else {
        $Filtered
      }
    }
    else {
      Write-Verbose -Message "Finding Tenant Dial Plan Names"
      Get-CsTenantDialPlan | Where-Object Identity -NE "Global" | Select-Object Identity
    }

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #Get-TeamsTDP