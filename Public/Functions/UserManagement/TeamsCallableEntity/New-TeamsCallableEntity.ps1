# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   Live




function New-TeamsCallableEntity {
  <#
  .SYNOPSIS
    Creates a Callable Entity for Auto Attendants
  .DESCRIPTION
    Wrapper for New-CsAutoAttendantCallableEntity with verification
    Requires a licensed User or ApplicationEndpoint an Office 365 Group or Tel URI
  .PARAMETER Identity
    Required. Tel URI, Group Name or UserPrincipalName, depending on the Entity Type
  .PARAMETER EnableTranscription
    Optional. Enables Transcription. Available only for Groups (Type SharedVoicemail)
  .PARAMETER Type
    Optional. Type of Callable Entity to create.
    Expected User, ExternalPstn, SharedVoicemail, ApplicationEndPoint
    If not provided, the Type is queried with Get-TeamsCallableEntity
  .PARAMETER ReturnObjectIdOnly
    Using this switch will return only the ObjectId of the validated CallableEntity, but will not create the Object
    This way the Command can be used to validate connected Objects for Call Queues.
  .PARAMETER Force
    Suppresses confirmation prompt to enable Users for Enterprise Voice, if required and $Confirm is TRUE
  .EXAMPLE
    New-TeamsAutoAttendantEntity -Type ExternalPstn -Identity "tel:+1555123456"
    Creates a callable Entity for the provided string, normalising it into a Tel URI
  .EXAMPLE
    New-TeamsAutoAttendantEntity -Type User -Identity John@domain.com
    Creates a callable Entity for the User John@domain.com
  .NOTES
    For Users, it will verify the Objects eligibility.
    Requires a valid license but can enable the User Object for Enterprise Voice if needed.
    For Groups, it will verify that the Group exists in AzureAd (but not in Exchange)
    For ExternalPstn it will construct the Tel URI
  .INPUTS
    System.String
  .OUTPUTS
    System.Object - (default)
  .COMPONENT
    TeamsAutoAttendant
    TeamsCallQueue
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
    New-TeamsAutoAttendantDialScope
	.LINK
    New-TeamsAutoAttendantMenu
	.LINK
    New-TeamsAutoAttendantPrompt
	.LINK
    New-TeamsAutoAttendantSchedule
  #>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
  [Alias('New-TeamsAutoAttendantCallableEntity', 'New-TeamsAAEntity')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, HelpMessage = 'Identity of the Call Target')]
    [string]$Identity,

    [Parameter(HelpMessage = 'Enables Transcription (for Shared Voicemail only)')]
    [switch]$EnableTranscription,

    [Parameter(HelpMessage = 'Callable Entity type: ExternalPstn, User, SharedVoiceMail, ApplicationEndpoint')]
    [ValidateSet('User', 'ExternalPstn', 'SharedVoicemail', 'ApplicationEndpoint')]
    [string]$Type,

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

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    # preparing Splatting Object
    $Parameters = $null

    # Normalising TelephoneNumber
    If ($Identity -match '^(tel:)?\+?(([0-9]( |-)?)?(\(?[0-9]{3}\)?)( |-)?([0-9]{3}( |-)?[0-9]{4})|([0-9]{7,15}))?((;( |-)?ext=[0-9]{3,8}))?$') {
      $Identity = Format-StringForUse $Identity -As E164 | Format-StringForUse -As LineURI
      Write-Verbose -Message "Callable Entity Type matches Phone Number - Number normalised to '$Identity'"
    }

    # Determining Callable Entity
    try {
      $CEObject = Get-TeamsCallableEntity "$Identity" -ErrorAction Stop
    }
    catch {
      Write-Error -Message "No Unique Target found for '$Identity'" -Exception System.Reflection.AmbiguousMatchException
      return
    }

    # Type
    if ( $Type ) {
      # Type is provided
      if ($CEObject.Type -ne $Type) {
        Write-Error -Message 'Callable Entity Type does not match queried type. Either omit the Type parameter or provide correct Type'
        return
      }
      else {
        Write-Verbose -Message 'Callable Entity Type matches queried type. OK'
      }
    }
    else {
      if ($CEObject.ObjectType -eq 'Unknown') {
        Write-Error -Message 'Object could not be determined and Cannot be used!' -ErrorAction Stop
      }
      else {
        # Determining Type
        Write-Verbose -Message "Callable Entity Type determined: '$($CEObject.Type)'"
      }
    }

    # Adding Parameters
    $Parameters = @{'Identity' = $CEObject.Identity }
    $Parameters += @{'Type' = $CEObject.Type }


    # EnableTranscription
    if ( $EnableTranscription ) {
      if ($CEObject.Type -eq 'SharedVoicemail') {
        Write-Information 'EnableTranscription - Transcription is activated for SharedVoicemail'
        $Parameters += @{'EnableTranscription' = $true }
      }
      else {
        Write-Verbose -Message 'EnableTranscription - Transcription can only be activated for SharedVoicemail.' -Verbose
      }
    }
    #endregion


    # Create CsAutoAttendantCallableEntity
    Write-Verbose -Message '[PROCESS] Creating Callable Entity'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess("$Identity", 'New-CsAutoAttendantCallableEntity')) {
      New-CsAutoAttendantCallableEntity @Parameters
      Write-Verbose -Message "$($MyInvocation.MyCommand) - created."
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsCallableEntity
