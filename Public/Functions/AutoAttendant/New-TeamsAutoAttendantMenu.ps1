# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  12-DEC-2020
# Status:   Live


#TODO Add output for Menu Name - to separate multiple executions

function New-TeamsAutoAttendantMenu {
  <#
  .SYNOPSIS
    Creates a Menu Object to be used in Auto Attendants
  .DESCRIPTION
    Creates a Menu Object with Prompt and/or MenuOptions to be used in Auto Attendants
    Wrapper for New-CsAutoAttendantMenu with friendly names
    Combines New-CsAutoAttendantMenu, New-CsAutoAttendantPrompt and New-CsAutoAttendantMenuOption
  .PARAMETER Name
    Optional. Name of the Menu if desired. Otherwise generated automatically.
  .PARAMETER Action
    Required. MenuOptions, Disconnect, TransferToCallTarget.
    Determines the type of Menu to be created.
  .PARAMETER Prompts
    Required for Action "MenuOptions" only. A Prompts Object, String or Full path to AudioFile.
    A Prompts Object will be used as is, otherwise it will be created dependent of the provided String
    A String will be used as Text-to-Voice. A File path ending in .wav, .mp3 or .wma will be used to create a recording.
  .PARAMETER MenuOptions
    Required for Action "MenuOptions" only. Mutually exclusive with CallTargetsInOrder.
    MenuOptions objects created with either New-TeamsAutoAttendantMenuOption or New-CsAutoAttenantMenuOption.
  .PARAMETER CallTargetsInOrder
    Required for Action "MenuOptions" only. Mutually exclusive with MenuOptions. Call Targets for Menu Options.
    Expected UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
    Allows to skip options with empty strings ("") or "$null". See Examples for details
  .PARAMETER CallTarget
    Required for Action "TransferToCallTarget" only. Single Call Target to redirect Calls to.
    UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
  .PARAMETER AddOperatorOnZero
    Optional for Action MenuOptions when used with CallTargetsInOrder only.
    This switch is ignored if more than nine (9) CallTargetsInOrder are specified
    Adds one more menu option to Transfer the Call to the Operator on pressing 0.
    NOTE: The AutoAttendant which will receive a Menu with this option, must have an Operator defined.
    Errors may occur if no operator is present.
  .PARAMETER EnableDialByName
    Required for Action "MenuOptions" only. Hashtable of Routing Targets.
    DTMF tones are assigned in order with 0 being TransferToOperator
  .PARAMETER DirectorySearchMethod
    Required for Action "MenuOptions" only. Hashtable of Routing Targets.
    DTMF tones are assigned in order with 0 being TransferToOperator
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Name "My Menu" -Action MenuOptions -Prompts $Prompts -MenuOptions $MenuOptions [-EnableDialByName] [-DirectorySearchMethod ByName]
    Classic behaviour, mostly synonymous with functionality provided by New-CsAutoAttendantMenu. Please see parameters there.
    Creates Menu with the MenuOptions Objects provided and applies the Prompts Object as the Greeting.
    Parameters EnableDialByName and DirectorySearchMethod can be used as outlined in New-CsAutoAttendantMenu
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action MenuOptions -Prompts "Press 1 for Sales..." -MenuOptions $MenuOptions -DirectorySearchMethod ByExtension
    Creates a Menu with a Prompt and MenuOptions. Creates a Prompts Object with the provided Text-to-voice string.
    Creates Menu with the MenuOptions Objects provided. DirectorySearchMethod is set to ByExtension
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action MenuOptions -Prompts "C:\temp\Menu.wav" -CallTargetsInOrder "MyCQ@domain.com","MyAA@domain.com","$null","tel:+15551234567"
    Creates a Menu with a Prompt and MenuOptions. Creates a Prompts Object with the provided Path to the Audio File.
    Creates a Menu Object with the Call Targets provided in order of application depending on identified ObjectType:
    Option 1 and 2 will be TransferToCallTarget (ApplicationEndpoint), Call Queue and Auto Attendant respectively.
    Option 3 will not be assigned ($null), Option 4 will be TransferToCallTarget (ExternalPstn)
    Maximum 10 options are supported, if more than 9 are provided, the Switch AddOperatorOnZero (if used) is ignored.
    NOTE: This method does not allow specifying User Object that are intended to forward to the Users Voicemail!
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action MenuOptions -Prompts "Press 1 for John, Press 3 for My Group" -CallTargetsInOrder "John@domain.com","","My Group"
    Creates a Menu with a Prompt and MenuOptions. Creates a Prompts Object with the provided Text-to-voice string.
    Creates a Menu with MenuOptions Objects provided in order with the Parameter CallTargetsInOrder depending on identified ObjectType:
    Option 1 will be TransferToCallTarget (User), Option 2 is unassigned (empty string),
    Option 3 is TransferToCallTarget (Shared Voicemail)
    Maximum 10 options are supported, if more than 9 are provided, the Switch AddOperatorOnZero (if used) is ignored.
    NOTE: This method does not allow specifying User Object that are intended to forward to the Users Voicemail!
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action Disconnect
    Creates a default Menu, disconnecting the Call
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action TransferToOperator
    Creates a default Menu, transferring the Call to the Operator
  .EXAMPLE
    New-TeamsAutoAttendantMenu -Action TransferToCallTarget -CallTarget "John@domain.com"
    Creates a default Menu, transferring the Call to the Call target.
    Expected UserPrincipalName (User, ApplicationEndpoint), Group Name (Shared Voicemail), Tel Uri (ExternalPstn)
  .NOTES
    Limitations: CallTargetsInOrder are Menu Options integrated and their type is parsed with Get-TeamsCallableEntity
    This provides the following limitations:
    1) Operator can only be specified as a redirection target on "Press 0" with Switch AddOperatorOnZero, not as an option in CallTargetsInOrder.
    2) Provided UPNs that are found to be AzureAdUsers are limited to be used as "Person in the Organisation"
    If forwarding to the Users Voicemail is required, please change this in the Admin Center afterwards.
    To overcome either limitation, please define MenuOptions yourself and use with the MenuOptions Parameter.

    To define Menu Options manually, please see:
    https://docs.microsoft.com/en-us/powershell/module/skype/new-csautoattendantmenuoption?view=skype-ps

    Please see 'Set up an auto attendant' for details:
    https://docs.microsoft.com/en-us/MicrosoftTeams/create-a-phone-system-auto-attendant?WT.mc_id=TeamsAdminCenterCSH
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
	.LINK
    https://docs.microsoft.com/en-us/MicrosoftTeams/create-a-phone-system-auto-attendant?WT.mc_id=TeamsAdminCenterCSH
	.LINK
    https://docs.microsoft.com/en-us/powershell/module/skype/new-csautoattendantmenuoption?view=skype-ps
  #>

  [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Disconnect', ConfirmImpact = 'Low')]
  [Alias('New-TeamsAAMenu')]
  [OutputType([System.Object])]
  param(
    [Parameter(HelpMessage = 'Optional Name of the Menu')]
    [ValidateLength(5, 63)]
    [string]$Name,

    [Parameter(Mandatory, HelpMessage = 'Action determines Type of Menu to be built')]
    [ValidateSet('TransferToMenu', 'Disconnect', 'TransferToCallTarget')]
    [string]$Action,

    [Parameter(Mandatory, ParameterSetName = 'MenuOptions', HelpMessage = 'Prompt object, Text-To-Voice String or Full path to AudioFile')]
    [Parameter(Mandatory, ParameterSetName = 'MenuOptions2', HelpMessage = 'Prompt object, Text-To-Voice String or Full path to AudioFile')]
    $Prompts,

    [Parameter(Mandatory, ParameterSetName = 'MenuOptions', HelpMessage = 'MenuOptions Object')]
    [object[]]$MenuOptions,

    [Parameter(Mandatory, ParameterSetName = 'MenuOptions2', HelpMessage = 'Up to 9 Call Targets in order to be applied as MenuOptions')]
    [AllowNull()]
    [AllowEmptyString()]
    [Alias('MenuOptionsInOrder')]
    [string[]]$CallTargetsInOrder,

    [Parameter(Mandatory, ParameterSetName = 'TransferToCallTarget', HelpMessage = 'Up to 9 Call Targets in order to be applied as MenuOptions')]
    [string]$CallTarget,

    [Parameter(ParameterSetName = 'MenuOptions2', HelpMessage = 'Adds a Menu Option for option 0 to Transfer to Operator')]
    [switch]$AddOperatorOnZero,

    [Parameter(HelpMessage = 'Enables directory search by recipient name and get transferred to the party')]
    [switch]$EnableDialByName,

    [Parameter(HelpMessage = 'Directory Search Method for the Auto Attendant menu')]
    [ValidateSet('None', 'ByName', 'ByExtension')]
    [string]$DirectorySearchMethod

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

    #region Routing - Menu Options
    switch ($Action) {
      'TransferToMenu' {
        #region Prompt
        $PromptsType = ($Prompts | Get-Member | Select-Object TypeName -First 1).TypeName
        switch ($PromptsType) {
          'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.Prompt' {
            Write-Verbose -Message 'Call Flow - Prompts provided is a Prompt Object'
            $Parameters += @{'Prompts' = $Prompts }

          }
          'System.String' {
            Write-Verbose -Message 'Call Flow - Greeting provided as a String'
            # Process Greeting
            try {
              $PromptsObject = New-TeamsAutoAttendantPrompt -String "$Prompts"
              if ($PromptsObject) {
                Write-Verbose -Message 'Prompts - Adding 1 Prompts created (Greeting)'
                $Parameters += @{'Prompts' = $PromptsObject }
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
        #endregion


        #region MenuOptions or CallTargetsInOrder
        switch ($PSCmdlet.ParameterSetName) {
          'MenuOptions' {
            # Determine Type
            <# This doesn't work - Type is Selected.System.Management.Automation.PSCustomObject
            foreach ($MenuOption in $MenuOptions) {
              $MenuOptionType = $null
              $MenuOptionType = ($MenuOption | Get-Member | Select-Object TypeName -First 1).TypeName
              if ($MenuOptionType -eq "Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.MenuOption") {
                Write-Verbose -Message "Menu Option - Provided Object is a Menu Object. Adding Menu Option"
              }
              else {
                Write-Error -Message "Menu Option - Provided Object not of correct Object Type. Please create a Menu with New-TeamsAutoAttendantMenuOption or New-CsAutoAttendantMenuOption" -ErrorAction Stop
              }
            }
            #>
          }

          'MenuOptions2' {
            # Process Ordered
            Write-Verbose -Message 'MenuOptions - Creating Menu Options for Call Targets as provided (CallTargetsInOrder)'
            $Option = 1
            $MaxOptions = 10
            #$MaxOptions = if ($AddOperatorOnZero) { 9 } else { 10 }
            if ($CallTargetsInOrder.Count -gt 10) {
              Write-Warning -Message 'MenuOptions - Max 10 options are supported as Call Targets. Additional Objects are ignored (CallTargetsInOrder)'
            }

            [System.Collections.ArrayList]$CreatedMenuOptions = @()
            foreach ($Target in $CallTargetsInOrder) {
              if ($Option -le $CallTargetsInOrder.Count -and $Option -le $MaxOptions) {
                if ( $Target ) {
                  $MenuOptionToAdd = $null
                  try {
                    if ( $Option -ne 10 ) {
                      $MenuOptionToAdd = New-TeamsAutoAttendantMenuOption -Press $Option -CallTarget $Target
                    }
                    else {
                      $MenuOptionToAdd = New-TeamsAutoAttendantMenuOption -Press 0 -CallTarget $Target
                    }

                    if ($MenuOptionToAdd) {
                      Write-Information "MenuOptions - Creating Option 'Press $Option' to '$Target' - OK"
                      [void]$CreatedMenuOptions.Add($MenuOptionToAdd)
                    }
                  }
                  catch {
                    Write-Warning -Message "MenuOptions - Creating Option 'Press $Option' to '$Target' - Creation unsuccessful! Omitting Call Target"
                  }
                }
                else {
                  Write-Verbose -Message "MenuOptions - Creating Option 'Press $Option' not provided, is empty or NULL - OK (omitted)" -Verbose
                }

                $Option++
              }
            }

            # AddOperatorOnZero
            if ($CallTargetsInOrder.Count -le 9 ) {
              if ($AddOperatorOnZero) {
                # Create Menu Option on "Press 0" to forward to Operator
                Write-Verbose -Message 'MenuOptions - Creating additional Option for Operator (AddOperatorOnZero)'
                $MenuOptionToAdd = $null
                $MenuOptionToAdd = New-TeamsAutoAttendantMenuOption -Press 0 -TransferToOperator
                if ($MenuOptionToAdd) {
                  Write-Verbose -Message 'SUCCESS (AddOperatorOnZero)'
                  [void]$CreatedMenuOptions.Add($MenuOptionToAdd)
                }
              }
            }
            else {
              Write-Verbose -Message 'MenuOptions - AddOperatorOnZero is not parsed as 10 Options were provided (CallTargetsInOrder)'
            }

            $MenuOptions = $CreatedMenuOptions
          }
        }
        #endregion

      } #MenuOption

      'Disconnect' {
        $MenuOptions = New-TeamsAutoAttendantMenuOption -DisconnectCall
      }

      'TransferToCallTarget' {
        # Change this to do MenuOptions
        $MenuOptions = New-TeamsAutoAttendantMenuOption -CallTarget $CallTarget
      }
    }

    # Adding MenuOptions
    Write-Verbose -Message "MenuOptions - Action '$Action' - Adding $($MenuOptions.Count) Menu Options"
    $Parameters += @{'MenuOptions' = $MenuOptions }
    #endregion


    #region Other Parameters
    if ( -not $Name) {
      $Name = "Menu with $($Parameters.MenuOptions.Count) Options" + $(if ($Parameters.Prompts) { ' and Greeting' })
    }
    $Parameters += @{'Name' = "$Name" }

    if ($EnableDialByName) {
      $Parameters += @{'EnableDialByName' = $true }
    }

    if ($DirectorySearchMethod) {
      $Parameters += @{'DirectorySearchMethod' = $DirectorySearchMethod }
    }
    #endregion


    # Create Menu
    Write-Verbose -Message '[PROCESS] Creating Menu'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess("$($Menu.Name)", 'New-CsAutoAttendantMenu')) {
      New-CsAutoAttendantMenu @Parameters
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsAutoAttendantMenu
