# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  12-DEC-2020
# Status:   Live




function New-TeamsAutoAttendantCallFlow {
  <#
  .SYNOPSIS
    Creates a Call Flow Object to be used in Auto Attendants
  .DESCRIPTION
    Creates a Call Flow with optional Prompt and Menu to be used in Auto Attendants
    Wrapper for New-CsAutoAttendantCallFlow with friendly names
    Combines New-CsAutoAttendantMenu, New-CsAutoAttendantPrompt
  .PARAMETER Name
    Optional. Name of the Call Flow if desired. Otherwise generated automatically.
  .PARAMETER Greeting
    Optional. A Prompts Object, String or Full path to AudioFile.
    A Prompts Object will be used as is, otherwise it will be created dependent of the provided String
    A String will be used as Text-to-Voice. A File ending in .wav, .mp3 or .wma will be used to create a recording.
  .PARAMETER Menu
    Optional. Menu Object to be used.
  .PARAMETER Disconnect
    Optional. Creates a default Menu, disconnecting the Call.
  .PARAMETER TransferToCallTarget
    Optional. String. Creates a default Menu, redirecting to the specified Call Target
    UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
  .EXAMPLE
    New-TeamsAutoAttendantCallFlow [-Name "Default Call Flow"] -Menu $MenuObject [-Greeting $PromptObject]
    Classic behaviour, synonymous with functionality provided by New-CsAutoAttendantCallFlow. Please see parameters there.
    Creates Call Flow with the Menu Object provided and optionally applies the PromptObject as the Greeting.
  .EXAMPLE
    New-TeamsAutoAttendantCallFlow -Menu $MenuObject -Greeting "Welcome to Contoso"
    Creates Call Flow with the Menu Object provided and creates the Greeting with the provided String (Text-to-voice)
  .EXAMPLE
    New-TeamsAutoAttendantCallFlow -TransferToCallTarget "John@domain.com"
    Creates a Menu Object to transfer the Call to a call Target and no Greeting
    UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
  .EXAMPLE
    New-TeamsAutoAttendantCallFlow -Disconnect
    Default. Creates Call Flow with a default Disconnect and no Greeting
  .NOTES
    Limitations: DialByName
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

  [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Disconnect', ConfirmImpact = 'Low')]
  [Alias('New-TeamsAAFlow')]
  [OutputType([System.Object])]
  param(
    [Parameter(HelpMessage = 'Optional Name of the Call Flow')]
    [ValidateLength(5, 63)]
    [string]$Name,

    [Parameter(HelpMessage = 'Prompt Object, Text-To-Voice String or Full path to AudioFile')]
    #Type is determined in BEGIN block
    $Greeting,

    [Parameter(ParameterSetName = 'Menu', HelpMessage = 'Menu Object to be used')]
    [Object]$Menu,

    [Parameter(ParameterSetName = 'Disconnect', HelpMessage = 'Creates a menu, using Disconnect')]
    [switch]$Disconnect,

    [Parameter(ParameterSetName = 'TransferToCallTarget', HelpMessage = 'Creates a menu, redirecting to the Call Target')]
    [string]$TransferToCallTarget
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

    #region Greeting
    if ($Greeting) {
      # Processing Greeting
      $GreetingType = ($Greeting | Get-Member | Select-Object TypeName -First 1).TypeName
      switch ($GreetingType) {
        'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.Prompt' {
          Write-Verbose -Message 'Call Flow - Greeting provided is a Prompt Object'
          $Parameters += @{'Greetings' = @($Greeting) }

        }
        'System.String' {
          Write-Verbose -Message 'Call Flow - Greeting provided as a String'
          # Process Greeting
          try {
            $GreetingObject = New-TeamsAutoAttendantPrompt -String "$Greeting"
            if ($GreetingObject) {
              Write-Verbose -Message 'Prompts - Adding 1 Prompts created (Greeting)'
              $Parameters += @{'Greetings' = $GreetingObject }
            }
          }
          catch {
            Write-Warning -Message "Call Flow - Menu - Greeting - Error creating prompt. Omitting Greeting. Exception Message: $($_.Exception.Message)"
          }
        }

        default {
          Write-Error -Message 'Type not accepted as a Greeting/Prompt, please provide a Prompts Object or a String' -ErrorAction Stop
        }
      }
    }
    #endregion


    #region Options
    # Processing Options
    switch ($PSCmdlet.ParameterSetName) {
      'Menu' {
        if ($Menu) {
          #<#
          $MenuType = ($Menu | Get-Member | Select-Object TypeName -First 1).TypeName
          if ($MenuType -eq 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.Menu') {
            Write-Verbose -Message 'Menu - Provided Object is a Menu Object. Adding Menu'
          }
          else {
            Write-Error -Message 'Menu - Provided Object not of correct Object Type. Please create a Menu with New-TeamsAutoAttendantMenu or New-CsAutoAttendantMenu' -ErrorAction Stop
          }
          #>
        }
        else {
          Write-Error -Message 'Menu - Provided Object is NULL' -ErrorAction Stop
        }
      }

      'Disconnect' {
        $Menu = New-TeamsAutoAttendantMenu -Action Disconnect
      }

      'TransferToCallTarget' {
        $Menu = New-TeamsAutoAttendantMenu -Action TransferToCallTarget -CallTarget $TransferToCallTarget
      }
    }

    $Parameters += @{'Menu' = $Menu }
    #endregion


    #region Other Parameters
    if ( -not $Name) {
      $Name = "Call Flow using '$($PSCmdlet.ParameterSetName)'" + $(if ($Parameters.Greetings) { ' and Greeting' })
    }
    $Parameters += @{'Name' = "$Name" }

    #endregion


    # Create Call Flow
    Write-Verbose -Message '[PROCESS] Creating Call Flow'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess("$($Parameters.Name)", 'New-CsAutoAttendantCallFlow')) {
      New-CsAutoAttendantCallFlow @Parameters
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsAutoAttendantCallFlow
