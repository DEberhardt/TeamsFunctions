# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   Live




function New-TeamsAutoAttendantDialScope {
  <#
  .SYNOPSIS
    Creates a Dial Scope to be used in Auto Attendants
  .DESCRIPTION
    Wrapper for New-CsAutoAttendantDialScope with friendly names
  .PARAMETER GroupName
    Required. Name of one or more Office 365 groups to create a Dial Scope for
  .EXAMPLE
    New-TeamsAutoAttendantDialScope -GroupName "My Group"
    Creates a Dial Scope for "My Group"
  .EXAMPLE
    New-TeamsAutoAttendantDialScope -GroupName "My Group","My other Group"
    Creates a Dial Scope including "My Group" and "My other Group"
  .INPUTS
    System.String
  .OUTPUTS
    System.Object
  .COMPONENT
    TeamsAutoAttendant
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
	.LINK
    New-TeamsAutoAttendant
	.LINK
    Set-TeamsAutoAttendant
	.LINK
    Get-TeamsCallableEntity
	.LINK
    Find-TeamsCallableEntity
	.LINK
    New-TeamsCallableEntity
	.LINK
    New-TeamsAutoAttendantCallFlow
	.LINK
    New-TeamsAutoAttendantMenu
	.LINK
    New-TeamsAutoAttendantMenuOption
	.LINK
    New-TeamsAutoAttendantPrompt
	.LINK
    New-TeamsAutoAttendantSchedule
	.LINK
    New-TeamsAutoAttendantDialScope
  #>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
  [Alias('New-TeamsAAScope')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory = $true, HelpMessage = 'Name of the Auto Attendant')]
    [string[]]$GroupName

  ) #param

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Asserting MicrosoftTeams Connection
    if (-not (Assert-MicrosoftTeamsConnection)) { break }

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference') }
    if (-not $PSBoundParameters.ContainsKey('WhatIf')) { $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    foreach ($Group in $GroupName) {
      Write-Verbose -Message "[PROCESS] Processing '$Group'"
      $Object = $null
      $Object = Get-TeamsCallableEntity -Identity "$Group"
      if ($Object) {

        $GroupIds += $Object.Identity
      }
      else {
        Write-Warning -Message "Group not found: '$Group' - Skipping"
      }

    }

    # Create dial Scope
    Write-Verbose -Message '[PROCESS] Creating Dial Scope'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      Write-Debug "$groupIds"
    }

    if ($PSCmdlet.ShouldProcess("$groupIds", 'New-CsAutoAttendantDialScope')) {
      New-CsAutoAttendantDialScope -GroupScope -GroupIds $groupIds
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsAutoAttendantDialScope
