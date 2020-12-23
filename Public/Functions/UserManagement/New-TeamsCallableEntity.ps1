﻿# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   RC




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
#    System.String - With Switch ReturnObjectIdOnly
  .COMPONENT
    TeamsAutoAttendant
    TeamsCallQueue
	.LINK
    New-TeamsAutoAttendant
    Set-TeamsAutoAttendant
    Get-TeamsCallableEntity
    Find-TeamsCallableEntity
    New-TeamsCallableEntity
    New-TeamsAutoAttendantDialScope
    New-TeamsAutoAttendantMenu
    New-TeamsAutoAttendantPrompt
    New-TeamsAutoAttendantSchedule
  #>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
  [Alias('New-TeamsAutoAttendantCallableEntity', 'New-TeamsAAEntity')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, Position = 0, HelpMessage = "Identity of the Call Target")]
    [string]$Identity,

    [Parameter(HelpMessage = "Enables Transcription (for Shared Voicemail only)")]
    [switch]$EnableTranscription,

    [Parameter(HelpMessage = "Callable Entity type: ExternalPstn, User, SharedVoiceMail, ApplicationEndpoint")]
    [ValidateSet('User', 'ExternalPstn', 'SharedVoicemail', 'ApplicationEndpoint')]
    [string]$Type,

    <#
    [Parameter(HelpMessage = "OutputType: Object or ObjectId")]
    [switch]$ReturnObjectIdOnly,
 #>
    [Parameter(HelpMessage = "Suppresses confirmation prompt to enable Users for Enterprise Voice, if Users are specified")]
    [switch]$Force

  ) #param

  begin {
    Show-FunctionStatus -Level RC
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Asserting SkypeOnline Connection
    if (-not (Assert-SkypeOnlineConnection)) { break }

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
      $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
    }
    if (-not $PSBoundParameters.ContainsKey('Confirm')) {
      $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference')
    }
    if (-not $PSBoundParameters.ContainsKey('WhatIf')) {
      $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference')
    }


  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    # preparing Splatting Object
    $Parameters = $null

    # Normalising TelephoneNumber
    If ($Identity -match "^(tel:)?\+?(([0-9]( |-)?)?(\(?[0-9]{3}\)?)( |-)?([0-9]{3}( |-)?[0-9]{4})|([0-9]{8,15}))?((;( |-)?ext=[0-9]{3,8}))?$") {
      $Identity = Format-StringForUse $Identity -As E164 | Format-StringForUse -As LineURI
      Write-Verbose -Message "Callable Entity Type matches Phone Number - Number normalised to '$Identity'" -Verbose
    }

    # Determining Callable Entity
    try {
      $CEObject = Get-TeamsCallableEntity $Identity -ErrorAction Stop
    }
    catch {
      Write-Error -Message "No Unique Target found for '$Identity'" -Exception System.Reflection.AmbiguousMatchException
      return
    }

    # Type
    if ( $Type ) {
      # Type is provided
      if ($CEObject.Type -ne $Type) {
        Write-Error -Message "Callable Entity Type does not match queried type. Either omit the Type parameter or provide correct Type"
        return
      }
      else {
        Write-Verbose -Message "Callable Entity Type matches queried type. OK"
      }
    }
    else {
      if ($CEObject.ObjectType -eq "Unknown") {
        Write-Error -Message "Object could not be determined and Cannot be used!" -ErrorAction Stop
        <# Code commented out as Choice is not a valid option when calling with another command
        #TODO Evaluate whether a handling parameter (silent?) can be introduced to error here instead of Giving a choice
        # Correcting Type if lookup fails
        $Title = "Type cannot be determined"
        $Prompt = "If the Object exists, please provide Type"
        $Choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&User", "&ExternalPstn", "&SharedVoicemail", "&ApplicationEndpoint", "E&xit")
        $Default = 5

        # Prompt for the choice
        $Type = $host.UI.PromptForChoice($Title, $Prompt, $Choices, $Default)

        if ($Type -eq "Exit") {
          return
        }
        #>
      }
      else {
        # Determining Type
        $Type = $CEObject.ObjectType
        Write-Verbose -Message "Callable Entity Type determined: '$Type'"
      }
    }

    # Adding Parameters
    $Parameters = @{'Identity' = $CEObject.Identity }
    $Parameters += @{'Type' = $Type }


    # EnableTranscription
    if ( $EnableTranscription ) {
      if ($CEObject.Type -eq "SharedVoicemail") {
        Write-Verbose -Message "EnableTranscription - Transcription is activated for SharedVoicemail"
        $Parameters += @{'EnableTranscription' = $true }
      }
      else {
        Write-Verbose -Message "EnableTranscription - Transcription can only be activated for SharedVoicemail." -Verbose
      }
    }
    #endregion


    # Create CsAutoAttendantCallableEntity
    Write-Verbose -Message "[PROCESS] Creating Callable Entity"
    if ($PSBoundParameters.ContainsKey('Debug')) {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess("$Identity", "New-CsAutoAttendantCallableEntity")) {
      New-CsAutoAttendantCallableEntity @Parameters
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsCallableEntity
