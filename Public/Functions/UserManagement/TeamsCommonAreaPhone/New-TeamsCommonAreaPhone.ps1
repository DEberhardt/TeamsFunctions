# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:		David Eberhardt
# Updated:  01-JAN-2021
# Status:   RC




function New-TeamsCommonAreaPhone {
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
    If not provided a Password will be generated with the string "CAP-" and todays date in the format: "CAP-03-JAN-2021"
	.PARAMETER IPPhonePolicy
		Optional. Adds an IP Phone Policy to the User
	.PARAMETER TeamsCallingPolicy
		Optional. Adds a Calling Policy to the User
	.PARAMETER TeamsCallParkPolicy
		Optional. Adds a Call Park Policy to the User
	.EXAMPLE
		New-TeamsCommonAreaPhone -UserPrincipalName "My Lobby Phone@TenantName.onmicrosoft.com" -UsageLocation US
		Will create a CommonAreaPhone with a Usage Location for 'US' and assign the CommonAreaPhone License
		User Principal Name will be normalised to: MyLobbyPhone@TenantName.onmicrosoft.com
		DisplayName will be taken from the User PrincipalName and normalised to "MyLobbyPhone"
    No Policies will be assigned to the Common Area Phone, the Global Policy will be in effect for this Phone
  .EXAMPLE
		New-TeamsCommonAreaPhone -UserPrincipalName "Lobby.@TenantName.onmicrosoft.com" -Displayname "Lobby {Phone}" -UsageLocation US -License CommonAreaPhone
		Will create a CommonAreaPhone with a Usage Location for 'US' and assign the CommonAreaPhone License
		User Principal Name will be normalised to: Lobby@TenantName.onmicrosoft.com
		DisplayName will be normalised to "Lobby Phone"
    No Policies will be assigned to the Common Area Phone, the Global Policy will be in effect for this Phone
  .EXAMPLE
		New-TeamsCommonAreaPhone -UserPrincipalName "Lobby@TenantName.onmicrosoft.com" -Displayname "Lobby Phone" -UsageLocation US -License Office365E3,PhoneSystem
    Will create a CommonAreaPhone with a Usage Location for 'US' and assign the Office 365 E3 License as well as PhoneSystem
    No Policies will be assigned to the Common Area Phone, the Global Policy will be in effect for this Phone
	.EXAMPLE
		New-TeamsCommonAreaPhone -UserPrincipalName "Lobby@TenantName.onmicrosoft.com" -Displayname "Lobby Phone" -UsageLocation US -IPPhonePolicy "My IPP" -TeamsCallingPolicy "CallP" -TeamsCallParkPolicy "CallPark"
		Will create a CommonAreaPhone with a Usage Location for 'US' and assign the CommonAreaPhone License
    The supplied Policies will be assigned to the Common Area Phone
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
		Creates a Common Area Phone in AzureAD for use in Teams
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
  [Alias('New-TeamsCAP')]
  [OutputType([System.Object])]
  param (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'UPN of the Object to create.')]
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

    [Parameter(Mandatory = $true, HelpMessage = 'Usage Location to assign')]
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

    [Parameter(HelpMessage = 'Password to be assigned to the account. Min 8 characters')]
    [ValidatePattern('^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)[a-zA-Z\d]{8,16}$')]
    [string]$Password,

    [Parameter(HelpMessage = 'IP Phone Policy')]
    [string]$IPPhonePolicy,

    [Parameter(HelpMessage = 'Teams Calling Policy')]
    [string]$TeamsCallingPolicy,

    [Parameter(HelpMessage = 'Teams Call Park Policy')]
    [string]$TeamsCallParkPolicy

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
    [int]$sMax = 9

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    $Parameters = @{}

    #region PREPARATION
    $Status = 'Verifying input'
    #region Normalising $UserPrincipalname
    $Operation = 'Normalising UserPrincipalName'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    $UPN = Format-StringForUse -InputString $UserPrincipalName -As UserPrincipalName
    Write-Verbose -Message "UserPrincipalName normalised to: '$UPN'"
    $Parameters += @{ 'UserPrincipalName' = "$UPN" }

    # MailNickName
    $MailnickName = Format-StringForUse -InputString $($UserPrincipalName.Split('@')[0]) -As DisplayName
    $Parameters += @{ 'MailNickName' = "$MailnickName" }
    #endregion

    #region Normalising $DisplayName
    $Operation = 'Normalising DisplayName'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    if ($PSBoundParameters.ContainsKey('DisplayName')) {
      $Name = Format-StringForUse -InputString $DisplayName -As DisplayName
    }
    else {
      $Name = $MailnickName
    }
    Write-Verbose -Message "DisplayName normalised to: '$Name'"
    $Parameters += @{ 'DisplayName' = "$Name" }
    #endregion

    #region UsageLocation
    $Operation = 'Parsing UsageLocation'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    if ($PSBoundParameters.ContainsKey('UsageLocation')) {
      Write-Verbose -Message "'$Name' UsageLocation parsed: Using '$UsageLocation'"
    }
    else {
      # Querying Tenant Country as basis for Usage Location
      # This is never triggered as UsageLocation is mandatory! Remaining here regardless
      $Tenant = Get-CsTenant -WarningAction SilentlyContinue
      if ($null -ne $Tenant.CountryAbbreviation) {
        $UsageLocation = $Tenant.CountryAbbreviation
        Write-Warning -Message "'$Name' UsageLocation not provided. Defaulting to: $UsageLocation. - Please verify and change if needed!"
      }
      else {
        Write-Error -Message "'$Name' Usage Location not provided and Country not found in the Tenant!" -Category ObjectNotFound -RecommendedAction 'Please run command again and specify -UsageLocation' -ErrorAction Stop
      }
    }
    $Parameters += @{ 'UsageLocation' = "$UsageLocation" }
    #endregion


    #region Password Profile
    $Operation = 'Password Profile'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.EnforceChangePasswordPolicy = $true
    $PasswordProfile.ForceChangePasswordNextLogin = $true
    if ($PSBoundParameters.ContainsKey('Password')) {
      $PasswordProfile.Password = $Password
    }
    else {
      $PasswordProfile.Password = 'CAP-' + $(Get-Date -Format 'dd-MMM-yyyy')
    }
    $Parameters += @{ 'PasswordProfile' = $PasswordProfile }
    $Parameters += @{ 'AccountEnabled' = $true }
    #endregion

    #Common Parameters
    $Parameters += @{ 'ErrorAction' = 'STOP' }
    #endregion


    #region ACTION
    $Status = 'Creating Object'
    #region Creating Account
    $Operation = 'Creating Common Area Phone'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    try {
      #Trying to create the Common Area Phone
      Write-Verbose -Message "'$Name' Creating Common Area Phone with New-AzureAdUser..."
      if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
        "Function: $($MyInvocation.MyCommand.Name) - Parameters", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
      }
      if ($PSCmdlet.ShouldProcess("$UPN", 'New-AzureAdUser')) {
        $AzureAdUser = New-AzureADUser @Parameters
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name) - AzureAdUser created", ($AzureAdUser | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }

        $i = 0
        $iMax = 20
        Write-Information "Common Area Phone '$Name' created; Waiting for AzureAd to write object ($iMax s)"
        $Status = 'Querying User'
        $Operation = 'Waiting for Get-AzureAdUser to return a Result'
        Write-Verbose -Message "$Status - $Operation"
        while ( -not (Test-AzureADUser $UPN)) {
          if ($i -gt $iMax) {
            Write-Error -Message "Could not find Object in AzureAD in the last $iMax Seconds" -Category ObjectNotFound -RecommendedAction 'Please verify Object has been created (UserPrincipalName); Continue with Set-TeamsCommonAreaPhone'
            return
          }
          Write-Progress -Id 1 -Activity 'Azure Active Directory is applying License. Please wait' `
            -Status $Status -SecondsRemaining $($iMax - $i) -CurrentOperation $Operation -PercentComplete (($i * 100) / $iMax)

          Start-Sleep -Milliseconds 1000
          $i++
        }
        Write-Progress -Id 1 -Activity 'Azure Active Directory is applying License. Please wait' -Status $Status -Completed
      }
      else {
        return
      }
    }
    catch {
      # Catching anything
      Write-Error "Common Area Phone '$Name' - Creation failed: $($_.Exception.Message)" -ForegroundColor Red
      return
    }
    #endregion

    $Status = 'Applying Settings'
    #region Licensing
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
    #endregion

    #region Policies
    $Operation = 'Applying Policies'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSBoundParameters.ContainsKey('IPPhonePolicy')) {
      Grant-CsTeamsIPPhonePolicy -Identity $AzureAdUser.ObjectId -PolicyName $IPPhonePolicy
    }
    else {
      Write-Verbose -Message 'No IP Phone Policy supplied - Global Policy is in effect!'
    }

    if ($PSBoundParameters.ContainsKey('TeamsCallingPolicy')) {
      Grant-CsTeamsCallingPolicy -Identity $AzureAdUser.ObjectId -PolicyName $TeamsCallingPolicy
    }
    else {
      Write-Verbose -Message 'No Calling Policy supplied - Global Policy is in effect!'
    }

    if ($PSBoundParameters.ContainsKey('TeamsCallParkPolicy')) {
      Grant-CsTeamsCallParkPolicy -Identity $AzureAdUser.ObjectId -PolicyName $TeamsCallParkPolicy
    }
    else {
      Write-Verbose -Message 'No Call Park Policy supplied - Global Policy is in effect!'
    }
    #endregion

    #region OUTPUT
    $Status = 'Validation'
    $Operation = 'Querying Object'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    $ObjectCreated = $null
    $ObjectCreated = Get-TeamsCommonAreaPhone -Identity $UPN
    #CHECK Output and Password application
    #$ObjectCreated | Add-Member -MemberType NoteProperty -Name Password -Value $AzureAdUser.Password
    $ObjectCreated | Add-Member -MemberType NoteProperty -Name Password -Value $PasswordProfile.Password
    Write-Output $ObjectCreated

    Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    #endregion

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsCommonAreaPhone
