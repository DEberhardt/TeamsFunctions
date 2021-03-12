# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  12-DEC-2020
# Status:   Live




function New-TeamsAutoAttendantMenuOption {
  <#
  .SYNOPSIS
    Creates a Menu Options Object
  .DESCRIPTION
    Creates a Menu Options Object to be used in Auto Attendants
    Wrapper for New-CsAutoAttendantMenuOption with friendly names
  .PARAMETER DisconnectCall
    Required to create a basic 'Disconnect' option. Switch. Default.
  .PARAMETER Press
    Required for Option TransferToOperator and TransferToCallTarget. Integer.
    Dtmf Tone (digit) to be pressed for this option
  .PARAMETER OrSay
    Optional for Option TransferToCallTarget. String.
    Voice Response to be used for this option. Expected: Single word
  .PARAMETER ToOperator
    Required for Option TransferToOperator. Switch. No other input necessary.
    NOTE: The AutoAttendant which will receive a Menu with this option, must have an Operator defined.
    Errors may occur if no operator is present.
  .PARAMETER ToCallTarget
    Required for Option TransferToCallTarget. String identifying the Call Target:
    UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Disconnect
    Creates a default Menu Option to be used for disconnecting the call.
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 0 -TransferToOperator
    Creates a Menu Option on pressing 0 (voice response is 'Operator' by default) to Transfer to the Operator.
    Note: The Operator must be specified in the AutoAttendant!
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 1 -CallTarget "My Group"
    Creates a Menu Option on pressing 1 or saying 'one' (default) to Transfer to the Call Target (Shared Voicemail)
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 2 -CallTarget Sales@domain.com -OrSay "Sales"
    Creates a Menu Option on pressing 2 or saying 'Sales' to Transfer to the Call Target (User).
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 3 -CallTarget MyCQ@domain.com -OrSay "Queue"
    Creates a Menu Option on pressing 3 or saying 'Queue' to Transfer to the Call Target (Call Queue).
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 4 -CallTarget MyAA@domain.com -OrSay "Menu"
    Creates a Menu Option on pressing 4 or saying 'Menu' to Transfer to the Call Target (Auto Attendant).
  .EXAMPLE
    New-TeamsAutoAttendantMenuOption -Press 5 -CallTarget "tel:+15551234567" -OrSay "Engineer"
    Creates a Menu Option on pressing 5 or saying 'Engineer' to Transfer to the Call Target (ExternalPstn).
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

  [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'DisconnectCall', ConfirmImpact = 'Low')]
  [Alias('New-TeamsAAOption')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, ParameterSetName = 'DisconnectCall', Position = 0, HelpMessage = 'Option to disconnect for default menus')]
    [switch]$DisconnectCall,

    [Parameter(ParameterSetName = 'Operator', HelpMessage = 'Number to press on the Dial Pad')]
    [Parameter(ParameterSetName = 'CallTarget', HelpMessage = 'Number to press on the Dial Pad')]
    [Alias('DtmfResponseTone')]
    [ValidateRange(0, 9)]
    [int]$Press,

    [Parameter(Mandatory, ParameterSetName = 'Operator', HelpMessage = 'Option to transfer the Call to the Operator defined')]
    [Alias('Operator')]
    [switch]$TransferToOperator,

    [Parameter(ParameterSetName = 'Operator', HelpMessage = 'Alternative voice Response')]
    [Parameter(ParameterSetName = 'CallTarget', HelpMessage = 'Alternative voice Response')]
    [Alias('VoiceResponses', 'Say')]
    [ValidateScript( { if ($_ -match '^[^\W_]+$') { $true } else { Write-Error -Message 'Voice Response must be one word without spaces or symbols.' } })]
    [string]$OrSay,

    [Parameter(ParameterSetName = 'CallTarget', HelpMessage = 'CallTarget')]
    [string]$CallTarget

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

    # Preparing Splatting Object
    $Parameters = $null
  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    switch ($PSCmdlet.ParameterSetName) {
      'DisconnectCall' {
        $Parameters += @{'DtmfResponse' = 'Automatic' }
        $Parameters += @{'Action' = 'DisconnectCall' }

      }

      'Operator' {
        if ($Press) {
          $DtmfResponse = 'Tone' + $Press
          if ($OrSay) {
            Write-Warning -Message "Parameter 'OrSay' can only be used together with 'Press' - omitted."
          }
        }
        else {
          $DtmfResponse = 'Automatic'
          if ($OrSay) {
            Write-Warning -Message "Parameter 'OrSay' can only be used together with 'Press' - omitted."
          }
        }

        $Parameters += @{'DtmfResponse' = $DtmfResponse }
        $Parameters += @{'Action' = 'TransferCallToOperator' }

      }

      'CallTarget' {
        if ($Press) {
          $DtmfResponse = 'Tone' + $Press
          if ($OrSay) {
            $Parameters += @{'VoiceResponses' = $OrSay }
          }
        }
        else {
          $DtmfResponse = 'Automatic'
          if ($OrSay) {
            Write-Warning -Message "Parameter 'OrSay' can only be used together with 'Press' - omitted."
          }
        }

        $Parameters += @{'DtmfResponse' = $DtmfResponse }
        $Parameters += @{'Action' = 'TransferCallToTarget' }


        # Determine Call Target
        try {
          $CallableEntity = New-TeamsCallableEntity -Identity "$CallTarget"
          if ( $CallableEntity ) {
            $Parameters += @{'CallTarget' = $CallableEntity }
          }
        }
        catch {
          Write-Error -Message "Error Creating Call Target: $($_.Exception.Message)" -ErrorAction Stop
        }
      }
    }

    # Create Menu Option
    Write-Verbose -Message '[PROCESS] Creating Menu Option'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess('New MenuOption', 'New-CsAutoAttendantMenuOption')) {
      New-CsAutoAttendantMenuOption @Parameters
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsAutoAttendantMenu
