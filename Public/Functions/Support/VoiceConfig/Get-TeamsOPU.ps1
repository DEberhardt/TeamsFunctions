﻿# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:		David Eberhardt
# Updated:  01-JAN-2021
# Status:   Live




function Get-TeamsOPU {
  <#
  .SYNOPSIS
    Lists all Online PSTN Usages by Name
  .DESCRIPTION
    To quickly find Online PSTN Usages, an Alias-Function to Get-CsOnlinePstnUsage
  .PARAMETER Usage
    String. Name or part of the Online Pstn Usage. Can be omitted to list Names of all Usages.
    Searches for Usages with Get-CsOnlinePstnUsage, listing all that match.
  .EXAMPLE
    Get-TeamsOPU
    Lists Identities (Names) of all Online Pstn Usages
  .EXAMPLE
    Get-TeamsOPU "PstnUsageName"
    Lists all PstnUsages with the String 'PstnUsageName' in the name of the Online Pstn Usage
  .NOTES
    This script is indulging the lazy admin. It behaves like (Get-CsOnlinePstnUsage).Usage
    This CmdLet behaves slightly different than the others, due to the nature of Pstn Usages.
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Get-TeamsOVP
  .LINK
    Get-TeamsOPU
  .LINK
    Get-TeamsOVR
  .LINK
    Get-TeamsMGW
  .LINK
    Get-TeamsTDP
  .LINK
    Get-TeamsVNR
  .LINK
    Get-TeamsIPP
  .LINK
    Get-TeamsCP
  .LINK
    Get-TeamsECP
  .LINK
    Get-TeamsECRP
  #>

  [CmdletBinding()]
  param (
    [Parameter(Position = 0, HelpMessage = 'Name of the Voice Routing Policy')]
    [string]$Usage
  )

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting MicrosoftTeams Connection
    if (-not (Assert-MicrosoftTeamsConnection)) { break }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    $Filtered = Get-CsOnlinePstnUsage Global
    if ($PSBoundParameters.ContainsKey('Usage')) {
      Write-Verbose -Message "Finding Online Pstn Usages with Usage '$Usage'"
      $Filtered = $Filtered | Where-Object Usage -Like "*$Usage*"
    }

    return $Filtered | Sort-Object Usage | Select-Object Usage -ExpandProperty Usage

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} # Get-TeamsOPU
