# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  01-DEC-2020
# Status:   Live


#IMPROVE? Evaluate better display: ToString manipulation. , to Line feed ; Reordering of objects (Menu Option: DtmfResponse, VoiceResponse, Action, Call Target)
#  Add TimeZone to main output (UTC+/-) and detailed output
function New-TeamsAutoAttendant {
  <#
	.SYNOPSIS
		Support function wrapping around New-CsAutoAttendant
	.DESCRIPTION
    This script handles select and limited variety for what Auto Attendants have to offer
    It should be seen as an extension rather than a replacement of New-CsAutoAttendant.
    It is currently still in development!
		UserPrincipalNames can be provided instead of IDs, FileNames (FullName) can be provided instead of IDs
	.PARAMETER Name
		Name of the Auto Attendant. Name will be normalised (unsuitable characters are filtered)
		Used as the DisplayName - Visible in Teams
  .PARAMETER TimeZone
    Required. TimeZone Identifier based on Get-CsAutoAttendantSupportedTimeZone, but abbreviated for easier input.
    Warning: Due to multiple time zone names with in the same relative difference to UTC this MAY produce incongruent output
    The time zone will be correct, but only specifying "UTC+01:00" for example will select the first entry.
    Default Value: "UTC"
  .PARAMETER LanguageId
    Required. Language Identifier indicating the language that is used to play text and identify voice prompts.
    Default Value: "en-US"
  .PARAMETER Operator
    Optional. Creates a Callable entity for the Operator
    Expected are UserPrincipalName (User, ApplicationEndPoint), a TelURI (ExternalPstn), an Office 365 Group Name (SharedVoicemail)
  .PARAMETER BusinessHoursGreeting
    Optional. Creates a Greeting for the Default Call Flow (during business hours) utilising New-TeamsAutoAttendantPrompt
    A supported Audio File or a text string that is parsed by the text-to-voice engine in the Language specified
    The last 4 digits will determine the type. For an AudioFile they are expected to be the file extension: '.wav', '.wma' or 'mp3'
    If DefaultCallFlow is provided, this parameter will be ignored.
  .PARAMETER BusinessHoursCallFlowOption
    Optional. Disconnect, TransferCallToTarget, Menu. Default is Disconnect.
    TransferCallToTarget requires BusinessHoursCallTarget. Menu requires BusinessHoursMenu
    If DefaultCallFlow is provided, this parameter will be ignored.
  .PARAMETER BusinessHoursCallTarget
    Optional. Requires BusinessHoursCallFlowOption to be TransferCallToTarget. Creates a Callable entity for this Call Target.
    Expected are UserPrincipalName (User, ApplicationEndPoint), a TelURI (ExternalPstn), an Office 365 Group Name (SharedVoicemail)
    If DefaultCallFlow is provided, this parameter will be ignored.
  .PARAMETER BusinessHoursMenu
    Optional. Requires BusinessHoursCallFlowOption to be Menu and a BusinessHoursCallTarget
    If DefaultCallFlow is provided, this parameter will be ignored.
  .PARAMETER AfterHoursGreeting
    Optional. Creates a Greeting for the After Hours Call Flow utilising New-TeamsAutoAttendantPrompt
    A supported Audio File or a text string that is parsed by the text-to-voice engine in the Language specified
    The last 4 digits will determine the type. For an AudioFile they are expected to be the file extension: '.wav', '.wma' or 'mp3'
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
  .PARAMETER AfterHoursCallFlowOption
    Optional. Disconnect, TransferCallToTarget, Menu. Default is Disconnect.
    TransferCallToTarget requires AfterHoursCallTarget. Menu requires AfterHoursMenu
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
  .PARAMETER AfterHoursCallTarget
    Optional. Requires AfterHoursCallFlowOption to be TransferCallToTarget. Creates a Callable entity for this Call Target
    Expected are UserPrincipalName (User, ApplicationEndPoint), a TelURI (ExternalPstn), an Office 365 Group Name (SharedVoicemail)
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
  .PARAMETER AfterHoursMenu
    Optional. Requires AfterHoursCallFlowOption to be Menu and a AfterHoursCallTarget
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
  .PARAMETER AfterHoursSchedule
    Optional. Default Schedule to apply: One of: MonToFri9to5 (default), MonToFri8to12and13to18, Open24x7
    A more granular Schedule can be used with the Parameter -Schedule
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
  .PARAMETER Schedule
    Optional. Custom Schedule object to apply for After Hours Call Flow
    Object created with New-TeamsAutoAttendantSchedule or New-CsAutoAttendantSchedule
    If CallFlows or CallHandlingAssociations are provided, this parameter will be ignored.
    Using this parameter to define the Schedule will override the Parameter -AfterHoursSchedule
  .PARAMETER EnableVoiceResponse
    Optional Switch to be passed to New-CsAutoAttendant
  .PARAMETER DefaultCallFlow
    Optional. Call Flow Object to pass to New-CsAutoAttendant (used as the Default Call Flow)
    Using this parameter to define the default Call Flow overrides all -BusinessHours Parameters
  .PARAMETER CallFlows
    Optional. Call Flow Object to pass to New-CsAutoAttendant
    Using this parameter to define additional Call Flows overrides all -AfterHours Parameters
    Requires Parameter CallHandlingAssociations in conjunction
  .PARAMETER CallHandlingAssociations
    Optional. Call Handling Associations Object to pass to New-CsAutoAttendant
    Using this parameter to define additional Call Flows overrides all -AfterHours Parameters
    Requires Parameter CallFlows in conjunction
  .PARAMETER InclusionScope
    Optional. DialScope Object to pass to New-CsAutoAttendant
    Object created with New-TeamsAutoAttendantDialScope or New-CsAutoAttendantDialScope
  .PARAMETER ExclusionScope
    Optional. DialScope Object to pass to New-CsAutoAttendant
    Object created with New-TeamsAutoAttendantDialScope or New-CsAutoAttendantDialScope
  .PARAMETER EnableTranscription
    Optional. Where possible, tries to enable Voicemail Transcription.
    Effective only for SharedVoicemail Targets as an Operator or MenuOption. Otherwise has no effect.
  .PARAMETER Force
    Suppresses confirmation prompt to enable Users for Enterprise Voice, if Users are specified
    Currently no other impact
	.EXAMPLE
		New-TeamsAutoAttendant -Name "My Auto Attendant"
    Creates a new Auto Attendant "My Auto Attendant" with Defaults
    TimeZone is UTC, Language is en-US and Schedule is Mon-Fri 9to5.
    Business hours and After Hours action is Disconnect
	.EXAMPLE
		New-TeamsAutoAttendant -Name "My Auto Attendant" -TimeZone UTC-05:00 -LanguageId pt-BR -AfterHoursSchedule MonToFri8to12and13to18 -EnableVoiceResponse
    Creates a new Auto Attendant "My Auto Attendant" and sets the TimeZone to UTC-5 and the language to Portuguese (Brazil)
    The Schedule of Mon-Fri 8to12 and 13to18 will be applied. Also enables VoiceResponses
	.EXAMPLE
		New-TeamsAutoAttendant -Name "My Auto Attendant" -Operator "tel:+1555123456"
    Creates a new Auto Attendant "My Auto Attendant" with default TimeZone and Language, but defines an Operator as a Callable Entity (Forward to Pstn)
	.EXAMPLE
    New-TeamsAutoAttendant -Name "My Auto Attendant" -BusinessHoursGreeting "Welcome to Contoso" -BusinessHoursCallFlowOption TransferCallToTarget -BusinessHoursCallTarget $CallTarget
    Creates a new Auto Attendant "My Auto Attendant" with defaults, but defines a Text-to-Voice Greeting, then forwards the Call to the Call Target.
    The CallTarget is queried based on input and created as required. UserPrincipalname for Users or ResourceAccount, Group Name for SharedVoicemail, provided as a string in the Variable $UPN
    This example is equally applicable to AfterHours.
	.EXAMPLE
		New-TeamsAutoAttendant -Name "My Auto Attendant" -DefaultCallFlow $DefaultCallFlow -CallFlows $CallFlows -InclusionScope $InGroups -ExclusionScope $OutGroups
    Creates a new Auto Attendant "My Auto Attendant" and passes through all objects provided. In this example, provided Objects are
    passed on through tto New-CsAutoAttendant and override other respective Parmeters provided:
    - A DefaultCallFlow Object is passed on which overrides all "-BusinessHours"-Parmeters
    - One or more CallFlows Objects are passed on which override all "-AfterHours"-Parameters
    - One or more CallHandlingAssociation Objects are passed on which override all "-AfterHours"-Parameters
    - An InclusionScope and an ExclusionScope are defined. These are passed on as-is
		All other values, like Language and TimeZone are defined with their defaults and can still be defined with the Objects.
  .INPUTS
    System.String
  .OUTPUTS
    System.Object
	.NOTES
		Currently in Testing
	.FUNCTIONALITY
		Creates a Auto Attendant with custom settings and friendly names as input
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
	.LINK
		New-TeamsCallQueue
	.LINK
    New-TeamsAutoAttendant
	.LINK
    Set-TeamsAutoAttendant
	.LINK
    Get-TeamsCallableEntity
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
    Remove-TeamsAutoAttendant
	.LINK
    New-TeamsResourceAccount
	.LINK
    New-TeamsResourceAccountAssociation
	#>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
  [Alias('New-TeamsAA')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, ValueFromPipeline, HelpMessage = 'Name of the Auto Attendant')]
    [string]$Name,

    [Parameter(HelpMessage = 'TimeZone Identifier')]
    [ValidateSet('UTC-12:00', 'UTC-11:00', 'UTC-10:00', 'UTC-09:00', 'UTC-08:00', 'UTC-07:00', 'UTC-06:00', 'UTC-05:00', 'UTC-04:30', 'UTC-04:00', 'UTC-03:30', 'UTC-03:00', 'UTC-02:00', 'UTC-01:00', 'UTC', 'UTC+01:00', 'UTC+02:00', 'UTC+03:00', 'UTC+03:30', 'UTC+04:00', 'UTC+04:30', 'UTC+05:00', 'UTC+05:30', 'UTC+05:45', 'UTC+06:00', 'UTC+06:30', 'UTC+07:00', 'UTC+08:00', 'UTC+09:00', 'UTC+09:30', 'UTC+10:00', 'UTC+11:00', 'UTC+12:00', 'UTC+13:00', 'UTC+14:00')]
    [string]$TimeZone = 'UTC',

    [Parameter(HelpMessage = 'Language Identifier from Get-CsAutoAttendantSupportedLanguage.')]
    [ValidateScript( { $_ -in (Get-CsAutoAttendantSupportedLanguage).Id })]
    [string]$LanguageId = 'en-US',

    [Parameter(Mandatory = $false, HelpMessage = 'Target String for the Operator (UPN, Group Name or Tel URI')]
    [string]$Operator,

    [Parameter(HelpMessage = 'Business Hours Greeting - Text String or Recording')]
    [string]$BusinessHoursGreeting,

    [Parameter(HelpMessage = 'Business Hours Call Flow - Default options')]
    [ValidateSet('Disconnect', 'TransferCallToTarget', 'Menu')]
    [string]$BusinessHoursCallFlowOption,

    [Parameter(HelpMessage = 'Business Hours Call Target - BusinessHoursCallFlowOption = TransferCallToTarget')]
    [string]$BusinessHoursCallTarget,

    [Parameter(HelpMessage = 'Business Hours Call Target - BusinessHoursCallFlowOption = Menu')]
    [object]$BusinessHoursMenu,

    [Parameter(HelpMessage = 'After Hours Greeting - Text String or Recording')]
    [string]$AfterHoursGreeting,

    [Parameter(HelpMessage = 'After Hours Call Flow - Default options')]
    [ValidateSet('Disconnect', 'TransferCallToTarget', 'Menu')]
    [string]$AfterHoursCallFlowOption,

    [Parameter(HelpMessage = 'After Hours Call Target - AfterHoursCallFlowOption = TransferCallToTarget')]
    [string]$AfterHoursCallTarget,

    [Parameter(HelpMessage = 'After Hours Call Target - AfterHoursCallFlowOption = Menu')]
    [object]$AfterHoursMenu,

    [Parameter(HelpMessage = 'Default Schedule to apply')]
    [ValidateSet('Open24x7', 'MonToFri9to5', 'MonToFri8to12and13to18')]
    [string]$AfterHoursSchedule,

    [Parameter(HelpMessage = 'Schedule Object created with New-TeamsAutoAttendantSchedule to apply')]
    [object]$Schedule,

    #Default Parameters of New-CsAutoAttendant for Pass-through application
    [Parameter(HelpMessage = 'Voice Responses')]
    [switch]$EnableVoiceResponse,

    [Parameter(HelpMessage = 'Default Call Flow')]
    [object]$DefaultCallFlow,

    [Parameter(HelpMessage = 'Call Flows')]
    [object]$CallFlows,

    [Parameter(HelpMessage = 'CallHandlingAssociations')]
    [object]$CallHandlingAssociations,

    [Parameter(HelpMessage = 'Groups defining the Inclusion Scope')]
    [object]$InclusionScope,

    [Parameter(HelpMessage = 'Groups defining the Exclusion Scope')]
    [object]$ExclusionScope,

    [Parameter(HelpMessage = 'Tries to Enable Transcription wherever possible')]
    [switch]$EnableTranscription,

    [Parameter(HelpMessage = 'Suppresses confirmation prompt to enable Users for Enterprise Voice, if Users are specified')]
    [switch]$Force

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

    # Initialising counters for Progress bars
    [int]$step = 0
    [int]$sMax = 8
    if ( -not $DefaultCallFlow ) {
      $sMax++
      if ( $BusinessHoursGreeting ) { $sMax++ }
    }
    if ( -not $CallFlows ) {
      $sMax++
      if ( $AfterHoursGreeting ) { $sMax++ }
      if ( -not $Schedule ) { $sMax++ }
    }

    #region Parameter validation
    $Status = 'Verifying input'
    $Operation = 'Validating Parameters'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    # Language has to be normalised as the Id is case sensitive. Default value: en-US
    $Language = $($LanguageId.Split('-')[0]).ToLower() + '-' + $($LanguageId.Split('-')[1]).ToUpper()
    Write-Verbose "LanguageId '$LanguageId' normalised to '$Language'"
    $VoiceResponsesSupported = (Get-CsAutoAttendantSupportedLanguage -Id $Language).VoiceResponseSupported

    # TimeZoneId - Generated from $TimeZone. Default value: UTC
    Write-Verbose -Message "TimeZone - Parsing TimeZone '$TimeZone'"
    if ($TimeZone -eq 'UTC') {
      $TimeZoneId = $TimeZone
    }
    else {
      $TimeZoneId = (Get-CsAutoAttendantSupportedTimeZone | Where-Object DisplayName -Like "($TimeZone)*" | Select-Object -First 1).Id
      Write-Verbose -Message "TimeZone - Found! Using: '$TimeZoneId'"
      Write-Information 'TimeZone - This is a correct match for the Time Zone, but might not be fully precise. - Please fine-tune Time Zone in the Admin Center if needed.'
    }

    #region BusinessHours
    # Main Call Flow -- DefaultCallFlow VS BusinessHours*
    if ($DefaultCallFlow) {
      # DefaultCallFlow
      Write-Information 'DefaultCallFlow - Overriding all BusinessHours-Parameters'

      if ($PSBoundParameters.ContainsKey('BusinessHoursGreeting')) { $PSBoundParameters.Remove('BusinessHoursGreeting') }
      if ($PSBoundParameters.ContainsKey('BusinessHoursCallFlowOption')) { $PSBoundParameters.Remove('BusinessHoursCallFlowOption') }
      if ($PSBoundParameters.ContainsKey('BusinessHoursCallTarget')) { $PSBoundParameters.Remove('BusinessHoursCallTarget') }

      # Testing provided Object Type
      if (($DefaultCallFlow | Get-Member | Select-Object TypeName -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.CallFlow') {
        Write-Error "DefaultCallFlow - Type is not of 'Microsoft.Rtc.Management.Hosted.OAA.Models.CallFlow'. Please provide a Call Flow Object" -Category InvalidType
        break
      }
    }
    else {
      # BusinessHours Parameters
      if (-not $PSBoundParameters.ContainsKey('BusinessHoursCallFlowOption')) {
        Write-Verbose -Message "BusinessHoursCallFlowOption - Parameter not specified. Defaulting to 'Disconnect' No other 'BusinessHours'-Parameters are processed!"
        $BusinessHoursCallFlowOption = 'Disconnect'
      }
      elseif ($BusinessHoursCallFlowOption -eq 'TransferCallToTarget') {
        # Must contain Target
        if (-not $PSBoundParameters.ContainsKey('BusinessHoursCallTarget')) {
          Write-Error -Message "BusinessHoursCallFlowOption (TransferCallToTarget) - Parameter 'BusinessHoursCallTarget' missing"
          break
        }

        # Must not contain a Menu
        if ($PSBoundParameters.ContainsKey('BusinessHoursMenu')) {
          Write-Verbose -Message 'BusinessHoursCallFlowOption (TransferCallToTarget) - Parameter BusinessHoursMenu cannot be used and will be omitted!' -Verbose
          $PSBoundParameters.Remove('BusinessHoursMenu')
        }
      }
      elseif ($BusinessHoursCallFlowOption -eq 'Menu') {
        # Must contain a Menu
        if (-not $PSBoundParameters.ContainsKey('BusinessHoursMenu')) {
          Write-Error -Message 'BusinessHoursCallFlowOption (Menu) - BusinessHoursMenu missing'
          break
        }
        else {
          # Testing provided Object Type
          if (($BusinessHoursMenu | Get-Member | Select-Object -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.Menu') {
            Write-Error -Message "BusinessHoursCallFlowOption (Menu) - BusinessHoursMenu not of the Type 'Microsoft.Rtc.Management.Hosted.OAA.Models.Menu'" -Category InvalidType
            break
          }
        }

        # Must not contain Target
        if ($PSBoundParameters.ContainsKey('BusinessHoursCallTarget')) {
          Write-Verbose -Message "BusinessHoursCallFlowOption (Menu) - Parameter 'BusinessHoursCallTarget' cannot be used and will be omitted!" -Verbose
          $PSBoundParameters.Remove('BusinessHoursCallTarget')
        }
      }
    }
    #endregion

    #region AfterHours
    # Call Flows & Call Handling Associations
    if ($CallFlows -or $CallHandlingAssociations) {
      # Custom Call Flows
      Write-Information 'CallFlows - Overriding all AfterHours-Parameters'
      if ($PSBoundParameters.ContainsKey('AfterHoursGreeting')) { $PSBoundParameters.Remove('AfterHoursGreeting') }
      if ($PSBoundParameters.ContainsKey('AfterHoursCallFlowOption')) { $PSBoundParameters.Remove('AfterHoursCallFlowOption') }
      if ($PSBoundParameters.ContainsKey('AfterHoursCallTarget')) { $PSBoundParameters.Remove('AfterHoursCallTarget') }
      if ($PSBoundParameters.ContainsKey('AfterHoursSchedule')) { $PSBoundParameters.Remove('AfterHoursSchedule') }
      if ($PSBoundParameters.ContainsKey('Schedule')) { $PSBoundParameters.Remove('Schedule') }


      if ($CallFlows -and -not $CallHandlingAssociations) {
        Write-Error -Message 'CallFlows - Parameter requires CallHandlingAssociation to be specified'
        break
      }

      if ($CallHandlingAssociations -and -not $CallFlows) {
        Write-Error -Message 'CallHandlingAssociations - Parameter requires CallFlows to be specified'
        break
      }

      # Testing provided Object Type
      foreach ($Flow in $CallFlows) {
        if (($Flow | Get-Member | Select-Object -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.CallFlow') {
          Write-Error -Message "CallFlows - '$($Flow.Name)' -Object not of the Type 'Microsoft.Rtc.Management.Hosted.OAA.Models.CallFlow'" -Category InvalidType
          break
        }
      }

      # Testing provided Object Type
      foreach ($CHA in $CallHandlingAssociations) {
        if (($CHA | Get-Member | Select-Object -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.CallHandlingAssociation') {
          Write-Error -Message "CallHandlingAssociations - '$($CHA.Name)' -Object not of the Type 'Microsoft.Rtc.Management.Hosted.OAA.Models.CallHandlingAssociation'" -Category InvalidType
          break
        }
      }
    }
    else {
      # AfterHours Parameters
      if (-not $PSBoundParameters.ContainsKey('AfterHoursCallFlowOption')) {
        Write-Warning -Message "AfterHoursCallFlowOption - Parameter not specified. Defaulting to 'Disconnect' No other 'BusinessHours'-Parameters are processed!"
        $AfterHoursCallFlowOption = 'Disconnect'
      }
      elseif ($AfterHoursCallFlowOption -eq 'TransferCallToTarget') {
        # Must contain Target
        if (-not $PSBoundParameters.ContainsKey('AfterHoursCallTarget')) {
          Write-Error -Message "AfterHoursCallFlowOption (TransferCallToTarget) - Parameter 'AfterHoursCallTarget' missing"
          break
        }

        # Must not contain a Menu
        if ($PSBoundParameters.ContainsKey('AfterHoursMenu')) {
          Write-Verbose -Message 'AfterHoursCallFlowOption (TransferCallToTarget) - Parameter AfterHoursMenu cannot be used and will be omitted!' -Verbose
          $PSBoundParameters.Remove('AfterHoursMenu')
        }
      }
      elseif ($AfterHoursCallFlowOption -eq 'Menu') {
        # Must contain a Menu
        if (-not $PSBoundParameters.ContainsKey('AfterHoursMenu')) {
          Write-Error -Message 'AfterHoursCallFlowOption (Menu) - AfterHoursMenu missing'
          break
        }
        else {
          if (($AfterHoursMenu | Get-Member | Select-Object -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.OAA.Models.Menu') {
            Write-Error -Message "AfterHoursCallFlowOption (Menu) - AfterHoursMenu not of the Type 'Microsoft.Rtc.Management.Hosted.OAA.Models.Menu'" -Category InvalidType
            break
          }
        }

        # Must not contain Target
        if ($PSBoundParameters.ContainsKey('AfterHoursCallTarget')) {
          Write-Verbose -Message "AfterHoursCallFlowOption (Menu) - Parameter 'AfterHoursCallTarget' cannot be used and will be omitted!"-Verbose
          $PSBoundParameters.Remove('AfterHoursCallTarget')
        }
      } # AfterHours Parameters

      #region Schedule & AfterHoursSchedule
      if ($Schedule) {
        if ($AfterHoursSchedule) {
          Write-Information 'Schedule - Custom Schedule Object overrides AfterHoursSchedule provided'
          $PSBoundParameters.Remove('AfterHoursSchedule')
        }

        # Testing provided Object Type
        if (($Schedule | Get-Member | Select-Object TypeName -First 1).TypeName -ne 'Deserialized.Microsoft.Rtc.Management.Hosted.Online.Models.Schedule') {
          Write-Error "Schedule - Type is not of 'Microsoft.Rtc.Management.Hosted.Online.Models.Schedule'. Please provide a Schedule Object" -Category InvalidType
          break
        }
      }
      else {
        if ( $AfterHoursSchedule) {
          Write-Information "Schedule - AfterHoursSchedule provided, Using: '$AfterHoursSchedule'"
        }
        else {
          $AfterHoursSchedule = 'MonToFri9to5'
        }

        # Creating Schedule
        $Operation = 'Creating Schedule'
        $step++
        Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
        Write-Verbose -Message "$Status - $Operation"

        $Schedule = switch ($AfterHoursSchedule) {
          'Open24x7' {
            New-TeamsAutoAttendantSchedule -Name 'Business Hours Schedule' -WeeklyRecurrentSchedule -BusinessDays MonToSun -BusinessHours AllDay -Complement
          }
          'MonToFri9to5' {
            New-TeamsAutoAttendantSchedule -Name 'Business Hours Schedule' -WeeklyRecurrentSchedule -BusinessDays MonToFri -BusinessHours 9to5 -Complement
          }
          'MonToFri8to12and13to18' {
            New-TeamsAutoAttendantSchedule -Name 'Business Hours Schedule' -WeeklyRecurrentSchedule -BusinessDays MonToFri -BusinessHours 8to12and13to18 -Complement
          }
        }
        Write-Verbose -Message "Schedule - Schedule created: '$AfterHoursSchedule'"
      }
    }
    #endregion

    #endregion
    #endregion

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    # re-Initialising counters for Progress bars
    [int]$step = 0

    #region PREPARATION
    $Status = 'Preparing Parameters'
    # preparing Splatting Object
    $Parameters = $null

    #region Required Parameters
    $Operation = 'Name, TimeZone & Language, Voice Responses'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    # Normalising $Name
    $NameNormalised = Format-StringForUse -InputString $Name -As DisplayName
    Write-Verbose -Message "'$Name' DisplayName normalised to: '$NameNormalised'"
    $Parameters += @{'Name' = $NameNormalised }

    # Preparing Call Flow String (to adhere to 64 Character limit)
    $CallFlowNamePrefix = -join "$NameNormalised"[0..40]
    if ($NameNormalised.length -gt 40) {
      Write-Verbose 'Auto Attendant Name is too long and cannot be used for Call Flow Name(s) as-is. Name will be shortened'
    }

    # Adding required parameters
    $Parameters += @{'LanguageId' = $Language }
    $Parameters += @{'TimeZoneId' = $TimeZoneId }

    # EnableVoiceResponse
    if ($PSBoundParameters.ContainsKey('EnableVoiceResponse')) {
      # Checking whether Voice Responses are available for the provided Language
      if ($VoiceResponsesSupported) {
        # Using As-Is
        Write-Verbose -Message "'$NameNormalised' EnableVoiceResponse - Voice Responses are supported with Language '$Language' and will be activated (Switch 'EnableVoiceResponse' will be used)"
        $Parameters += @{'EnableVoiceResponse' = $true }
      }
      else {
        Write-Warning -Message "'$NameNormalised' EnableVoiceResponse - Voice Responses are not supported for Language '$Language' and cannot be activated (Switch 'EnableVoiceResponse' will be omitted)"
      }
    }
    #endregion

    #region Operator
    $Operation = 'Operator'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSBoundParameters.ContainsKey('Operator')) {
      try {
        $OperatorEntity = New-TeamsCallableEntity -Identity $Operator
        if ($OperatorEntity) {
          $Parameters += @{'Operator' = $OperatorEntity }
        }
      }
      catch {
        Write-Warning -Message 'Operator - Error creating Call Target - skipped'
      }
    }
    #endregion


    #region Business Hours Call Flow
    $Operation = 'Business Hours Call Flow - Default Call Flow & Call Flow Option'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ( $DefaultCallFlow ) {
      # Using As-Is
      Write-Information "'$NameNormalised' DefaultCallFlow - Custom Object provided. Over-riding other options (like switch 'BusinessHoursCallFlow')"
      $Parameters += @{'DefaultCallFlow' = $DefaultCallFlow }

    }
    else {
      Write-Verbose -Message "'$NameNormalised' DefaultCallFlow - No Custom Object - Processing 'BusinessHoursCallFlowOption'..."
      $BusinessHoursCallFlowParameters = @{}
      $BusinessHoursCallFlowParameters.Name = "$NameNormalised - Business Hours CF"

      #region Processing BusinessHoursCallFlowOption
      switch ($BusinessHoursCallFlowOption) {
        'TransferCallToTarget' {
          Write-Verbose -Message "'$NameNormalised' DefaultCallFlow - Transferring to Target"

          # Process BusinessHoursCallTarget
          try {
            $BusinessHoursCallTargetEntity = New-TeamsCallableEntity "$BusinessHoursCallTarget" -ErrorAction Stop

            # Building Menu Only if Successful
            if ($BusinessHoursCallTargetEntity) {
              $BusinessHoursMenuOptionTransfer = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $BusinessHoursCallTargetEntity -DtmfResponse Automatic
              $BusinessHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($BusinessHoursMenuOptionTransfer)
              Write-Information "'$NameNormalised' Business Hours Call Flow - Menu (TransferCallToTarget) created"
              break
            }
            else {
              # Reverting to Disconnect
              Write-Warning -Message "'$NameNormalised' DefaultCallFlow - Business Hours Menu not created properly. Reverting to Disconnect"
              $BusinessHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
              $BusinessHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($BusinessHoursMenuOptionDefault)
            }
          }
          catch {
            Write-Warning -Message 'BusinessHoursCallTarget - Error creating Call Target - Defaulting to disconnect'
            $BusinessHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $BusinessHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($BusinessHoursMenuOptionDefault)
          }
        }

        'Menu' {
          Write-Verbose -Message "'$NameNormalised' DefaultCallFlow - Menu"
          if ($PSBoundParameters.ContainsKey('BusinessHoursMenu')) {
            # Menu is passed on as-is - $BusinessHoursMenu is defined and attached
            $BusinessHoursMenuObject = $BusinessHoursMenu
            Write-Information "'$NameNormalised' Business Hours Call Flow - Menu (BusinessHoursMenu) used"
          }
          else {
            # No custom / default Menu is currently created
            # $BusinessHoursMenu is Mandatory. If this is built out, the check against this must also be removed!
          }
        }

        default {
          # Defaulting to Disconnect
          Write-Verbose -Message "'$NameNormalised' DefaultCallFlow not provided or 'Disconnect' - Using default (Disconnect)"
          $BusinessHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
          $BusinessHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($BusinessHoursMenuOptionDefault)
        }
      }
      #endregion

      #region BusinessHoursGreeting
      #Adding optional BusinessHoursGreeting
      if ($PSBoundParameters.ContainsKey('BusinessHoursGreeting')) {
        $Operation = 'Business Hours Call Flow - Greeting'
        $step++
        Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
        Write-Verbose -Message "$Status - $Operation"

        try {
          $BusinessHoursGreetingObject = New-TeamsAutoAttendantPrompt -String "$BusinessHoursGreeting"
          if ($BusinessHoursGreetingObject) {
            Write-Information "'$NameNormalised' Business Hours Call Flow - Greeting created"
            $BusinessHoursCallFlowParameters.Greetings = @($BusinessHoursGreetingObject)
          }
        }
        catch {
          Write-Warning -Message "'$NameNormalised' CallFlow - BusinessHoursCallFlow - Greeting not enumerated. Omitting Greeting"
        }
      }
      #endregion

      #region Building Call Flow
      $Operation = 'Business Hours Call Flow - Building Call Flow'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Adding Business Hours Call Flow
      $BusinessHoursCallFlowParameters.Menu = $BusinessHoursMenuObject
      $BusinessHoursCallFlow = New-CsAutoAttendantCallFlow @BusinessHoursCallFlowParameters
      Write-Information "'$NameNormalised' Business Hours Call Flow - Call Flow created"
      $Parameters += @{'DefaultCallFlow' = $BusinessHoursCallFlow }
      #endregion
    }
    #endregion

    #region After Hours (Call Flow, Schedule & Call Handling Association)
    #region After Hours Call Flow
    #Initialising Variables for Call Handling Association
    $AfterHoursCallHandlingAssociationParams = @{}
    $AfterHoursCallHandlingAssociationParams.Type = 'AfterHours'

    # Processing CallFlow
    $Operation = 'After Hours Call Flow - Call Flows & Call Flow Option'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSBoundParameters.ContainsKey('CallFlows')) {
      # Custom Option provided - Using As-Is
      Write-Information "'$NameNormalised' CallFlow - Custom Object provided. Over-riding other options (like switch 'AfterHoursCallFlow')"
      $Parameters += @{'CallFlows' = $CallFlows }
    }
    else {
      # Option Selected
      Write-Verbose -Message "'$NameNormalised' CallFlow - No Custom Object - Processing 'AfterHoursCallFlowOption'..."
      $AfterHoursCallFlowParameters = @{}
      $AfterHoursCallFlowParameters.Name = "$CallFlowNamePrefix - After Hours CF"

      #region Processing AfterHoursCallFlowOption
      switch ($AfterHoursCallFlowOption) {
        'TransferCallToTarget' {
          Write-Verbose -Message "'$NameNormalised' CallFlow - Transferring to Target"

          # Process AfterHoursCallTarget
          try {
            $AfterHoursCallTargetEntity = New-TeamsCallableEntity "$AfterHoursCallTarget" -ErrorAction Stop

            # Building Menu Only if Successful
            if ($AfterHoursCallTargetEntity) {
              $AfterHoursMenuOptionTransfer = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $AfterHoursCallTargetEntity -DtmfResponse Automatic
              $AfterHoursMenuObject = New-CsAutoAttendantMenu -Name 'After Hours Menu' -MenuOptions @($AfterHoursMenuOptionTransfer)
              Write-Information "'$NameNormalised' After Hours Call Flow - Menu (TransferCallToTarget) created"
              break
            }
            else {
              # Reverting to Disconnect
              Write-Warning -Message "'$NameNormalised' Call Flow - After Hours Menu not created properly. Reverting to Disconnect"
              $AfterHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
              $AfterHoursMenuObject = New-CsAutoAttendantMenu -Name 'After Hours Menu' -MenuOptions @($AfterHoursMenuOptionDefault)
            }
          }
          catch {
            Write-Warning -Message 'AfterHoursCallTarget - Error creating Call Target - Defaulting to disconnect'
            $AfterHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $AfterHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($AfterHoursMenuOptionDefault)
          }
        }

        'Menu' {
          Write-Verbose -Message "'$NameNormalised' CallFlow - AfterHoursCallFlow - Menu"
          if ($PSBoundParameters.ContainsKey('AfterHoursMenu')) {
            # Menu is passed on as-is - $AfterHoursMenu is defined and attached
            $AfterHoursMenuObject = $AfterHoursMenu
            Write-Information "'$NameNormalised' After Hours Call Flow - Menu (BusinessHoursMenu) used"
          }
          else {
            # No custom / default Menu is currently created
            # $AfterHoursMenu is Mandatory. If this is built out, the check against this must also be removed!
          }
        }

        default {
          # Defaulting to Disconnect
          Write-Verbose -Message "'$NameNormalised' CallFlow - AfterHoursCallFlow not provided or Disconnect. Using default (Disconnect)"
          $AfterHoursMenuOptionDefault = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
          $AfterHoursMenuObject = New-CsAutoAttendantMenu -Name 'Business Hours Menu' -MenuOptions @($AfterHoursMenuOptionDefault)
        }
      }
      #endregion

      #region AfterHoursGreeting
      # Adding AfterHoursGreeting
      if ($PSBoundParameters.ContainsKey('AfterHoursGreeting')) {
        $Operation = 'After Hours Call Flow - Greeting'
        $step++
        Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
        Write-Verbose -Message "$Status - $Operation"

        try {
          $AfterHoursGreetingObject = New-TeamsAutoAttendantPrompt -String "$AfterHoursGreeting"
          if ($AfterHoursGreetingObject) {
            Write-Information "'$NameNormalised' After Hours Call Flow - Greeting created"
            $AfterHoursCallFlowParameters.Greetings = @($AfterHoursGreetingObject)
          }
        }
        catch {
          Write-Warning -Message "'$NameNormalised' CallFlow - AfterHoursCallFlow - Greeting not enumerated. Omitting Greeting"
        }
      }
      #endregion

      #region Building Call Flow
      $Operation = 'After Hours Call Flow - Building Call Flow'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Adding After Hours Call Flow
      $AfterHoursCallFlowParameters.Menu = $AfterHoursMenuObject
      $AfterHoursCallFlow = New-CsAutoAttendantCallFlow @AfterHoursCallFlowParameters
      Write-Information "'$NameNormalised' After Hours Call Flow - Call Flow created"
      $Parameters += @{'CallFlows' = $AfterHoursCallFlow }

      #TODO when HolidaySet is added, this needs to be array-proof (see processing of CallFlows Objects for code samples)
      #$AfterHoursCallHandlingAssociationParams.CallFlowId = $AfterHoursCallFlow.Id # This works, but want to try whether arraying works too
      $AfterHoursCallHandlingAssociationParams.CallFlowId += $AfterHoursCallFlow.Id
      #endregion

      #region After Hours Schedule & Call Handling Association
      $Operation = 'Schedule & Call Handling Association'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      $AfterHoursCallHandlingAssociationParams.ScheduleId = $Schedule.Id
      $AfterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation @AfterHoursCallHandlingAssociationParams
      #TODO when HolidaySet is added, a second CHA will need to be added here! +=?
      Write-Information "'$NameNormalised' After Hours Call Flow - Call Handling Association created with Schedule"
      $Parameters += @{'CallHandlingAssociation' = @($AfterHoursCallHandlingAssociation) }
      #endregion
    }
    #endregion
    #endregion


    #region Inclusion and Exclusion Scope
    $Operation = 'Dial Scopes - Inclusion and Exclusion Scope'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    # Inclusion Scope
    if ($PSBoundParameters.ContainsKey('InclusionScope')) {
      Write-Verbose -Message "'$NameNormalised' InclusionScope provided. Using as-is"
      $Parameters += @{'InclusionScope' = $InclusionScope }

    }
    else {
      #Scope is optional
      Write-Verbose -Message "'$NameNormalised' InclusionScope not defined. To create one, please run New-TeamsAutoAttendantDialScope or New-CsAutoAttendantDialScope"
    }

    # Exclusion Scope
    if ($PSBoundParameters.ContainsKey('ExclusionScope')) {
      Write-Verbose -Message "'$NameNormalised' ExclusionScope provided. Using as-is"
      $Parameters += @{'ExclusionScope' = $ExclusionScope }

    }
    else {
      #Scope is optional
      Write-Verbose -Message "'$NameNormalised' ExclusionScope not defined. To create one, please run New-TeamsAutoAttendantDialScope or New-CsAutoAttendantDialScope"
    }
    #endregion

    #region Common parameters
    $Parameters += @{'WarningAction' = 'Continue' }
    $Parameters += @{'ErrorAction' = 'Stop' }
    #endregion
    #endregion


    #region ACTION
    Write-Verbose -Message '[PROCESS] Creating Auto Attendant'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    # Create AA (New-CsAutoAttendant)
    $Status = 'Creating Object'
    $Operation = "Creating Auto Attendant: '$NameNormalised'"
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSCmdlet.ShouldProcess("$NameNormalised", 'New-CsAutoAttendant')) {
      try {
        # Create the Auto Attendant with all enumerated Parameters passed through splatting
        $null = (New-CsAutoAttendant @Parameters)
        Write-Information "Auto Attendant '$NameNormalised' created with all Parameters"
      }
      catch {
        Write-Error -Message "Error creating the Auto Attendant: $($_.Exception.Message)" -Category InvalidResult
        return
      }
    }
    else {
      return
    }
    #endregion


    #region OUTPUT
    $Status = 'Creating Object'
    $Operation = 'Querying Object'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    $AAFinal = Get-TeamsAutoAttendant -Name "$NameNormalised" -WarningAction SilentlyContinue
    Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    Write-Output $AAFinal
    #endregion

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"

  } #end
} #New-TeamsAutoAttendant
