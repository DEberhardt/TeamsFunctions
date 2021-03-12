# Module:   TeamsFunctions
# Function: AutoAttendant
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   Live




function New-TeamsAutoAttendantPrompt {
  <#
  .SYNOPSIS
    Creates a prompt
  .DESCRIPTION
    Wrapper for New-CsAutoAttendantPrompt for easier use
  .PARAMETER String
    Required. String as a Path for a Recording or a Greeting (Text-to-Voice)
  .EXAMPLE
    New-TeamsAutoAttendantPrompt -String "Welcome to Contoso"
    Creates a Text-to-Voice Prompt for the String
    Warning: This will break if the String ends in a supported File extension
  .EXAMPLE
    New-TeamsAutoAttendantPrompt -String "myAudioFile.mp3"
    Verifies the file exists, then imports it (with Import-TeamsAudioFile)
    Creates a Audio File Prompt after import.
  .NOTES
    Warning: This will break if the String ends in a supported File extension (WAV, WMA or MP3)
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
  [Alias('New-TeamsAAPrompt')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory = $true, HelpMessage = 'Path the the recording OR Text-to-Voice string')]
    [string]$String

  ) #param

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

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
    $Prompt = $null

    if ($String -match '.wav' -or $String -match '.wma' -or $String -match '.mp3') {
      #Recording
      if (-not (Test-Path $String)) {
        Write-Error -Message 'Auto Attendant Prompt - AudioFile - File not found.' -ErrorAction Stop
      }
      else {
        Write-Verbose -Message '[PROCESS] Creating Auto Attendant Prompt - AudioFile'
        if ($PSCmdlet.ShouldProcess("$Prompt", 'New-CsAutoAttendantPrompt')) {
          try {
            $audioFile = Import-TeamsAudioFile -ApplicationType AutoAttendant -File "$String"
            $Parameters += @{'ActiveType' = 'AudioFile' }
            $Parameters += @{'AudioFilePrompt' = "$audioFile" }
          }
          catch {
            Write-Error -Message "Importing Audio File failed: $($_.Exception.Message)" -ErrorAction Stop
          }
        }
      }
    }
    else {
      #Assume it is Text-to-speech
      Write-Verbose -Message '[PROCESS] Creating Auto Attendant Prompt - Text-to-Speech'
      if ($PSCmdlet.ShouldProcess("$Prompt", 'New-CsAutoAttendantPrompt')) {
        $Parameters += @{'ActiveType' = 'TextToSpeech' }
        $Parameters += @{'TextToSpeechPrompt' = "$String" }
      }
    }

    # Creating Prompt
    Write-Verbose -Message '[PROCESS] Creating Prompt'
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
    }

    if ($PSCmdlet.ShouldProcess("$Name", 'New-CsAutoAttendantPrompt')) {
      New-CsAutoAttendantPrompt @Parameters
    }
  }

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsAutoAttendantPrompt
