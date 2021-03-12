# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:		David Eberhardt
# Updated:  01-JAN-2021
# Status:   RC




function Set-TeamsCommonAreaPhone {
  <#
	.SYNOPSIS
		Creates a new Common Area Phone
	.DESCRIPTION
		Teams Call Queues and Auto Attendants require a Common Area Phone.
		It can carry a license and optionally also a phone number.
		This Function was designed to create the ApplicationInstance in AD,
		apply a UsageLocation to the corresponding AzureAD User,
		license the User and subsequently apply a phone number, all with one Command.
	.PARAMETER UserPrincipalName
		Required. The UPN for the new CommonAreaPhone. Invalid characters are stripped from the provided string
	.PARAMETER DisplayName
		Optional. The Name it will show up as in Teams. Invalid characters are stripped from the provided string
	.PARAMETER UsageLocation
		Required. Two Digit Country Code of the Location of the entity. Should correspond to the Phone Number.
		Before a License can be assigned, the account needs a Usage Location populated.
	.PARAMETER License
		Optional. Specifies the License to be assigned: PhoneSystem or PhoneSystem_VirtualUser
		If not provided, will default to PhoneSystem_VirtualUser
		Unlicensed Objects can exist, but cannot be assigned a phone number
    NOTE: PhoneSystem is an add-on license and cannot be assigned on its own. it has therefore been deactivated for now.
  .PARAMETER Password
    Optional. String. 8 to 16 characters, at least one uppercase letter, one lowercase letter and one number.
    If not provided a random Password is chosen and displayed in the Output
	.PARAMETER IPPhonePolicy
		Optional. Adds an IP Phone Policy to the User
	.PARAMETER TeamsCallingPolicy
		Optional. Adds a Calling Policy to the User
	.PARAMETER TeamsCallParkPolicy
		Optional. Adds a Call Park Policy to the User
	.PARAMETER PassThru
		Optional. Displays the Object after execution.
	.EXAMPLE
    Set-TeamsCommonAreaPhone -UserPrincipalName MyLobbyPhone@TenantName.onmicrosoft.com -Displayname "Lobby {Phone}"
    Changes the Object MyLobbyPhone@TenantName.onmicrosoft.com. DisplayName will be normalised to "Lobby Phone" and applied.
	.EXAMPLE
    Set-TeamsCommonAreaPhone -UserPrincipalName MyLobbyPhone@TenantName.onmicrosoft.com -UsageLocation US -License CommonAreaPhone
    Changes the Object MyLobbyPhone@TenantName.onmicrosoft.com. Usage Location is set to 'US' and the CommonAreaPhone License is assigned.
	.EXAMPLE
    Set-TeamsCommonAreaPhone -UserPrincipalName MyLobbyPhone@TenantName.onmicrosoft.com -License Office365E3,PhoneSystem
    Changes the Object MyLobbyPhone@TenantName.onmicrosoft.com. Usage Location is required to be set. Assigns the Office 365 E3 License as well as PhoneSystem
	.EXAMPLE
		Set-TeamsCommonAreaPhone -UserPrincipalName "MyLobbyPhone@TenantName.onmicrosoft.com" -IPPhonePolicy "My IPP" -TeamsCallingPolicy "CallP" -TeamsCallParkPolicy "CallPark" -PassThru
    Applies IPPhonePolicy, TeamsCallingPolicy and TeamsCallParkPolicy to the Common Area Phone
    Displays the Common Area Phone Object afterwards
  .INPUTS
    System.String
  .OUTPUTS
    System.Object
	.NOTES
    Execution requires User Admin Role in Azure AD
    To assign a Phone Number to this Object, please apply a full Voice Configuration using Set-TeamsUserVoiceConfig
    This includes Phone Number and Calling Plan or Online Voice Routing Policy and optionally a Tenant Dial Plan.
    This Script only covers relevant elements for Common Area Phones themselves.
	.FUNCTIONALITY
		Changes a Common Area Phone in AzureAD for use in Teams
  .COMPONENT
    TeamsUserVoiceConfig
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
	.LINK
    Get-TeamsCommonAreaPhone
	.LINK
    New-TeamsCommonAreaPhone
	.LINK
    Set-TeamsCommonAreaPhone
	.LINK
    Remove-TeamsCommonAreaPhone
	.LINK
    Find-TeamsUserVoiceConfig
	.LINK
    Get-TeamsUserVoiceConfig
	.LINK
    Set-TeamsUserVoiceConfig
  #>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
  [Alias('Set-TeamsCAP')]
  [OutputType([System.Void])]
  param (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'UPN of the Object to query.')]
    [ValidateScript( {
        If ($_ -match '@') {
          $True
        }
        else {
          Write-Host 'Must be a valid UPN' -ForegroundColor Red
          $false
        }
      })]
    [Alias('Identity')]
    [string]$UserPrincipalName,

    [Parameter(ValueFromPipelineByPropertyName, HelpMessage = 'Display Name for this Object')]
    [string]$DisplayName,

    [Parameter(HelpMessage = 'Usage Location to assign')]
    [string]$UsageLocation,

    [Parameter(HelpMessage = 'License to be assigned')]
    [ValidateScript( {
        $LicenseParams = (Get-AzureAdLicense).ParameterName.Split('', [System.StringSplitOptions]::RemoveEmptyEntries)
        if ($_ -in $LicenseParams) {
          return $true
        }
        else {
          Write-Host "Parameter 'License' - Invalid license string. Supported Parameternames can be found with Get-AzureAdLicense" -ForegroundColor Red
          return $false
        }
      })]
    [string]$License,

    [Parameter(HelpMessage = 'IP Phone Policy')]
    [string]$IPPhonePolicy,

    [Parameter(HelpMessage = 'Teams Calling Policy')]
    [string]$TeamsCallingPolicy,

    [Parameter(HelpMessage = 'Teams Call Park Policy')]
    [string]$TeamsCallParkPolicy,

    [Parameter(HelpMessage = 'No output is written by default, Switch PassThru will return changed object')]
    [switch]$PassThru

  ) #param

  begin {
    Show-FunctionStatus -Level RC
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
    [int]$sMax = 2
    if ( $DisplayName ) { $sMax++ }
    if ( $UsageLocation ) { $sMax++ }
    if ( $License ) { $sMax++ }
    if ( $PassThru ) { $sMax++ }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    $Parameters = @{}

    #region PREPARATION
    $Status = 'Verifying input'
    #region Lookup of UserPrincipalName
    $Operation = 'Querying Object'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    try {
      #Trying to query the Resource Account
      $CsOnlineUser = (Get-CsOnlineUser -Identity "$UserPrincipalName" -WarningAction SilentlyContinue -ErrorAction STOP)
      $CurrentDisplayName = $CsOnlineUser.DisplayName
      Write-Verbose -Message "'$UserPrincipalName' Teams Object found: '$CurrentDisplayName'"

      $Parameters += @{ 'ObjectId' = $CsOnlineUser.ObjectId }
    }
    catch {
      # Catching anything
      Write-Error -Message "'$UserPrincipalName' Teams Object not found!" -Category ObjectNotFound -RecommendedAction 'Please provide a valid UserPrincipalName of an existing Resource Account' #-ErrorAction Stop
      return
    }
    #endregion

    #region Normalising $DisplayName
    if ($PSBoundParameters.ContainsKey('DisplayName')) {
      $Operation = 'Normalising DisplayName'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      $Name = Format-StringForUse -InputString $DisplayName -As DisplayName
      Write-Verbose -Message "DisplayName normalised to: '$Name'"
      $Parameters += @{ 'DisplayName' = "$Name" }
    }
    else {
      $Name = $CurrentDisplayName
    }
    #endregion

    #region UsageLocation
    $CurrentUsageLocation = $CsOnlineUser.UsageLocation
    if ($PSBoundParameters.ContainsKey('UsageLocation')) {
      $Operation = 'Parsing UsageLocation'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      if ($Usagelocation -eq $CurrentUsageLocation) {
        Write-Verbose -Message "'$Name' Usage Location already set to: $CurrentUsageLocation"
      }
      else {
        Write-Verbose -Message "'$Name' Usage Location will be set to: $Usagelocation"
        $Parameters += @{ 'UsageLocation' = "$UsageLocation" }
      }
    }
    else {
      if ($null -ne $CurrentUsageLocation) {
        Write-Verbose -Message "'$Name' Usage Location currently set to: $CurrentUsageLocation"
      }
      else {
        if ($PSBoundParameters.ContainsKey('License')) {
          Write-Error -Message "'$Name' Usage Location not set!" -Category ObjectNotFound -RecommendedAction 'Please run command again and specify -UsageLocation'# -ErrorAction Stop
          return
        }
        else {
          Write-Warning -Message "'$Name' Usage Location not set! This is a requirement for License assignment and Phone Number"
        }
      }
    }
    #endregion

    #Common Parameters
    $Parameters += @{ 'ErrorAction' = 'STOP' }
    #endregion


    #region ACTION
    $Status = 'Azure Ad User'
    #region Setting Object
    $Operation = 'Applying Settings'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    try {
      #Trying to create the Common Area Phone
      Write-Verbose -Message "'$Name' Creating Common Area Phone with New-AzureAdUser..."
      if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
        "Function: $($MyInvocation.MyCommand.Name) - Parameters", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
      }
      if ($PSCmdlet.ShouldProcess("$UPN", 'Set-AzureAdUser')) {
        $null = Set-AzureADUser @Parameters
        $AzureAdUser = Get-AzureADUser -ObjectId $Parameters.ObjectId
      }
      else {
        return
      }
    }
    catch {
      # Catching anything
      Write-Host "ERROR:   Application of settings failed: $($_.Exception.Message)" -ForegroundColor Red
      return
    }
    #endregion

    #region Licensing
    if ($PSBoundParameters.ContainsKey('License')) {
      # Verifying License is available to be assigned
      # Determining available Licenses from Tenant
      $Operation = 'Querying Tenant Licenses'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"
      $TenantLicenses = Get-TeamsTenantLicense

      # Setting License to Common Area Phone if not provided
      if ( -not $PSBoundParameters.ContainsKey('License')) {
        $License = 'CommonAreaPhone'
      }

      # Verifying License is available
      $Operation = 'Verifying License is available'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"
      if ($License -eq 'CommonAreaPhone') {
        $RemainingCAPLicenses = ($TenantLicenses | Where-Object { $_.SkuPartNumber -eq 'MCOCAP' }).Remaining
        Write-Verbose -Message "INFO: $RemainingCAPLicenses Common Area Phone Licenses still available"
        if ($RemainingCAPLicenses -lt 1) {
          Write-Error -Message 'ERROR: No free PhoneSystem Virtual User License remaining in the Tenant.' -ErrorAction Stop
        }
        else {
          try {
            if ($PSCmdlet.ShouldProcess("$UPN", 'Set-TeamsUserLicense -Add CommonAreaPhone')) {
              $null = (Set-TeamsUserLicense -Identity $UPN -Add $License -ErrorAction STOP)
              Write-Information "'$Name' License assignment - '$License' SUCCESS"
            }
          }
          catch {
            Write-Error -Message "'$Name' License assignment failed for '$License'"
          }
        }
      }
      else {
        try {
          if ($PSCmdlet.ShouldProcess("$UPN", "Set-TeamsUserLicense -Add $License")) {
            $null = (Set-TeamsUserLicense -Identity $UPN -Add $License -ErrorAction STOP)
            Write-Information "'$Name' License assignment - '$License' SUCCESS"
          }
        }
        catch {
          Write-Error -Message "'$Name' License assignment failed for '$License'"
        }
      }
    }
    #endregion

    #region Policies
    $Operation = 'Applying Policies'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSBoundParameters.ContainsKey('IPPhonePolicy')) {
      Grant-CsTeamsIPPhonePolicy -Identity $AzureAdUser.ObjectId -PolicyName $IPPhonePolicy
    }

    if ($PSBoundParameters.ContainsKey('TeamsCallingPolicy')) {
      Grant-CsTeamsCallingPolicy -Identity $AzureAdUser.ObjectId -PolicyName $TeamsCallingPolicy
    }

    if ($PSBoundParameters.ContainsKey('TeamsCallParkPolicy')) {
      Grant-CsTeamsCallParkPolicy -Identity $AzureAdUser.ObjectId -PolicyName $TeamsCallParkPolicy
    }
    #endregion
    #endregion

    #region OUTPUT
    if ($PassThru) {
      $Status = 'Validation'
      $Operation = 'Querying Object'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      $CommonAreaPhone = $null
      $CommonAreaPhone = Get-TeamsCommonAreaPhone -Identity $UPN
      Write-Output $CommonAreaPhone
    }

    Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    #endregion

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsCommonAreaPhone
