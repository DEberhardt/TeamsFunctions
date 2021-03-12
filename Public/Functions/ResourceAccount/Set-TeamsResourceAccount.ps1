# Module:   TeamsFunctions
# Function: ResourceAccount
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   RC

#CHECK "Karen wants to call the manager"^^


function Set-TeamsResourceAccount {
  <#
	.SYNOPSIS
		Changes a new Resource Account
	.DESCRIPTION
		This function allows you to update Resource accounts for Teams Call Queues and Auto Attendants.
		It can carry a license and optionally also a phone number.
		This Function was designed to service the ApplicationInstance in AD,
		the corresponding AzureAD User and its license and enable use of a phone number, all with one Command.
	.PARAMETER UserPrincipalName
		Required. Identifies the Object being changed
	.PARAMETER DisplayName
		Optional. The Name it will show up as in Teams. Invalid characters are stripped from the provided string
	.PARAMETER ApplicationType
		CallQueue or AutoAttendant. Determines the association the account can have:
		A resource Account of the type "CallQueue" can only be associated with to a Call Queue
		A resource Account of the type "AutoAttendant" can only be associated with an Auto Attendant
		NOTE: Though switching the account type is possible, this is currently untested: Handle with Care!
	.PARAMETER UsageLocation
		Two Digit Country Code of the Location of the entity. Should correspond to the Phone Number.
		Before a License can be assigned, the account needs a Usage Location populated.
	.PARAMETER License
		Specifies the License to be assigned: PhoneSystem or PhoneSystem_VirtualUser
		If not provided, will default to PhoneSystem_VirtualUser
		Unlicensed Objects can exist, but cannot be assigned a phone number
		If a license already exists, it will try to swap the license to the specified one.
    NOTE: PhoneSystem is an add-on license and cannot be assigned on its own. it has therefore been deactivated for now.
	.PARAMETER PhoneNumber
		Changes the Phone Number of the object.
		Can either be a Microsoft Number or a Direct Routing Number.
		Requires the Resource Account to be licensed correctly
		Required format is E.164, starting with a '+' and 10-15 digits long.
  .PARAMETER PassThru
    By default, no output is generated, PassThru will display the Object changed
  .EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName ResourceAccount@TenantName.onmicrosoft.com -Displayname "My {ResourceAccount}"
		Will normalize the Display Name (i.E. remove special characters), then set it as "My ResourceAccount"
	.EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName AA-Mainline@TenantName.onmicrosoft.com -UsageLocation US
		Sets the UsageLocation for the Account in AzureAD to US.
	.EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName AA-Mainline@TenantName.onmicrosoft.com -License PhoneSystem_VirtualUser
		Requires the Account to have a UsageLocation populated. Applies the License to Resource Account AA-Mainline.
		If no license is assigned, will try to assign. If the license is already applied, no action is taken.
		NOTE: Swapping licenses is currently not possible.
	.EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName AA-Mainline@TenantName.onmicrosoft.com -PhoneNumber +1555123456
		Changes the Phone number of the Object. Will cleanly remove the Phone Number first before reapplying it.
		This will only succeed if the object is licensed correctly!
	.EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName AA-Mainline@TenantName.onmicrosoft.com -PhoneNumber $Null
		Removes the Phone number from the Object
	.EXAMPLE
		Set-TeamsResourceAccount -UserPrincipalName MyRessourceAccount@TenantName.onmicrosoft.com -ApplicationType AutoAttendant
		Switches MyResourceAccount to the Type AutoAttendant
		NOTE: This is currently untested, errors might occur simply because not all caveats could be captured.
		Handle with care!
  .INPUTS
    System.String
  .OUTPUTS
    None
	.FUNCTIONALITY
		Changes a resource Account in AzureAD for use in Teams
  .COMPONENT
    TeamsAutoAttendant
    TeamsCallQueue
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
	.LINK
    Get-TeamsResourceAccountAssociation
	.LINK
    New-TeamsResourceAccountAssociation
	.LINK
		Remove-TeamsResourceAccountAssociation
	.LINK
    New-TeamsResourceAccount
	.LINK
    Get-TeamsResourceAccount
	.LINK
    Find-TeamsResourceAccount
	.LINK
    Set-TeamsResourceAccount
	.LINK
    Remove-TeamsResourceAccount
	#>

  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
  [Alias('Set-TeamsRA')]
  [OutputType([System.Void])]
  param (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'UPN of the Object to change')]
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

    [Parameter(HelpMessage = 'Display Name is shown in Teams')]
    [string]$DisplayName,

    [Parameter(HelpMessage = 'CallQueue or AutoAttendant')]
    [ValidateSet('CallQueue', 'AutoAttendant', 'CQ', 'AA')]
    [Alias('Type')]
    [string]$ApplicationType,

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

    [Parameter(HelpMessage = 'Telephone Number to assign')]
    [Alias('Tel', 'Number', 'TelephoneNumber')]
    [AllowNull()]
    [AllowEmptyString()]
    [string]$PhoneNumber,

    [Parameter(HelpMessage = 'By default, no output is generated, PassThru will display the Object changed')]
    [switch]$PassThru
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
    [int]$sMax = 5
    if ( $DisplayName ) { $sMax = $sMax + 2 }
    if ( $UsageLocation ) { $sMax++ }
    if ( $ApplicationType ) { $sMax = $sMax + 2 }
    if ( $License ) { $sMax = $sMax + 2 }
    if ( $License -and $PhoneNumber ) { $sMax++ }
    if ( $PhoneNumber ) { $sMax++ }
    if ( $PassThru ) { $sMax++ }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    #region PREPARATION
    $Status = 'Verifying input'
    #region Lookup of UserPrincipalName
    $Operation = 'Querying Object'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    try {
      #Trying to query the Resource Account
      $Object = (Get-CsOnlineApplicationInstance -Identity $UserPrincipalName -WarningAction SilentlyContinue -ErrorAction STOP)
      $CurrentDisplayName = $Object.DisplayName
      Write-Verbose -Message "'$UserPrincipalName' OnlineApplicationInstance found: '$CurrentDisplayName'"
    }
    catch {
      # Catching anything
      Write-Error -Message "'$UserPrincipalName' OnlineApplicationInstance not found!" -Category ObjectNotFound -RecommendedAction 'Please provide a valid UserPrincipalName of an existing Resource Account' #-ErrorAction Stop
      return
    }
    #endregion

    #region Normalising $DisplayName
    if ($PSBoundParameters.ContainsKey('DisplayName')) {
      $Operation = 'DisplayName'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      $DisplayNameNormalised = Format-StringForUse -InputString $DisplayName -As DisplayName
      $Name = $DisplayNameNormalised
      Write-Verbose -Message "DisplayName normalised to: '$Name'"
    }
    else {
      $Name = $CurrentDisplayName
    }
    #endregion

    #region ApplicationType and Associations
    if ($PSBoundParameters.ContainsKey('ApplicationType')) {
      $Operation = 'Application Type'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Translating $ApplicationType (Name) to ID used by Commands.
      $AppId = GetAppIdFromApplicationType $ApplicationType
      $CurrentAppId = $Object.ApplicationId
      # Does the ApplicationType differ? Does it have to be changed?
      if ($AppId -eq $CurrentAppId) {
        # Application IDs match - Type does not need to be changed
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Application Type already set to: $ApplicationType"
      }
      else {
        # Finding all Associations to of this Resource Account to Call Queues or Auto Attendants
        $Associations = Get-CsOnlineApplicationInstanceAssociation -Identity $UserPrincipalName -WarningAction SilentlyContinue -ErrorAction Ignore
        if ($Associations.count -gt 0) {
          # Associations found. Aborting
          Write-Error -Message "'$Name ($UserPrincipalName )' ApplicationType cannot be changed! Object is associated with Call Queue or AutoAttendant." -Category OperationStopped -RecommendedAction 'Remove Associations with Remove-CsOnlineApplicationInstanceAssociation manually' -ErrorAction Stop
        }
        else {
          Write-Verbose -Message "'$Name ($UserPrincipalName )' Application Type will be changed to: $ApplicationType"
        }
      }
    }
    #endregion

    #region PhoneNumber
    $Operation = 'Phone Number'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    # Querying CurrentPhoneNumber
    try {
      $CurrentPhoneNumber = $Object.PhoneNumber.Replace('tel:', '')
      Write-Verbose -Message "'$Name ($UserPrincipalName )' Phone Number assigned currently: $CurrentPhoneNumber"
    }
    catch {
      $CurrentPhoneNumber = $null
      Write-Verbose -Message "'$Name ($UserPrincipalName )' Phone Number assigned currently: NONE"
    }

    if ($PSBoundParameters.ContainsKey('PhoneNumber')) {
      #Validating Phone Number
      #VALIDATE application of empty or Null does indeed remove the phonenumber!
      if ($PhoneNumber -eq '' -or $null -eq $PhoneNumber) {
        if ($CurrentPhoneNumber) {
          Write-Warning -Message "PhoneNumber is NULL or Empty. The Existing Number '$CurrentPhoneNumber' will be removed"
        }
        else {
          Write-Verbose -Message 'PhoneNumber is NULL or Empty, but no Number is currently assigned. No Action taken'
        }
        $PhoneNumber = $null
      }
      elseif ($PhoneNumber -match '^(tel:)?\+?(([0-9]( |-)?)?(\(?[0-9]{3}\)?)( |-)?([0-9]{3}( |-)?[0-9]{4})|([0-9]{7,15}))?((;( |-)?ext=[0-9]{3,8}))?$') {
        if ( $PhoneNumber -match 'ext' ) {
          Write-Warning -Message "PhoneNumber '$PhoneNumber' has an extension set. Resource Accounts do not allow applications of Extensions!"
        }
        $E164Number = Format-StringForUse $PhoneNumber -As E164
        #TEST Application of already assigned number
        if ($CurrentPhoneNumber -eq $E164Number) {
          Write-Verbose -Message "PhoneNumber '$E164Number' is already applied"
        }
        else {
          Write-Verbose -Message "PhoneNumber '$E164Number' is in a usable format and will be applied"
          # Checking number is free
          Write-Verbose -Message 'PhoneNumber - Finding Number assignments'
          $UserWithThisNumber = Find-TeamsUserVoiceConfig -PhoneNumber $E164Number
          if ($UserWithThisNumber) {
            Write-Error -Message "'$Name ($UserPrincipalName )' Number '$E164Number' is already assigned to another User" -Category NotImplemented -RecommendedAction 'Please specify a different Number ' -ErrorAction Stop
          }
        }
      }
      else {
        Write-Error -Message "PhoneNumber '$PhoneNumber' - Not a valid Phone number. Please provide a number starting with a + and 10 to 15 digits long" -ErrorAction Stop
      }
    }
    #endregion

    #region UsageLocation
    $Operation = 'Usage Location'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    $CurrentUsageLocation = (Get-AzureADUser -ObjectId "$UserPrincipalName" -WarningAction SilentlyContinue).UsageLocation
    if ($PSBoundParameters.ContainsKey('UsageLocation')) {
      if ($Usagelocation -eq $CurrentUsageLocation) {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Usage Location already set to: $CurrentUsageLocation"
      }
      else {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Usage Location will be set to: $Usagelocation"
      }
    }
    else {
      if ($null -ne $CurrentUsageLocation) {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Usage Location currently set to: $CurrentUsageLocation"
        $UsageLocation = $CurrentUsageLocation
      }
      else {
        if (($PSBoundParameters.ContainsKey('License')) -or ($PSBoundParameters.ContainsKey('PhoneNumber'))) {
          Write-Error -Message "'$Name ($UserPrincipalName )' Usage Location not set!" -Category ObjectNotFound -RecommendedAction 'Please run command again and specify -UsageLocation'# -ErrorAction Stop
          return
        }
        else {
          Write-Warning -Message "'$Name ($UserPrincipalName )' Usage Location not set! This is a requirement for License assignment and Phone Number"
        }
      }
    }
    #endregion

    #region Current License
    $Operation = 'License Assignment'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    if ($PSBoundParameters.ContainsKey('License') -or $PSBoundParameters.ContainsKey('PhoneNumber')) {
      $CurrentLicense = $null
      # Determining license Status of Object
      if (Test-TeamsUserLicense -Identity $UserPrincipalName -License PhoneSystem) {
        $CurrentLicense = 'PhoneSystem'
      }
      elseif (Test-TeamsUserLicense -Identity $UserPrincipalName -License PhoneSystemVirtualUser) {
        $CurrentLicense = 'PhoneSystemVirtualUser'
      }
      if ($null -ne $CurrentLicense) {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Current License assigned: $CurrentLicense"
      }
      else {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' Current License assigned: NONE"
      }
    }
    #endregion
    #endregion


    #region ACTION
    $Status = 'Applying Settings'
    #region DisplayName
    if ($PSBoundParameters.ContainsKey('DisplayName')) {
      $Operation = 'DisplayName'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      try {
        if ($PSCmdlet.ShouldProcess("$UserPrincipalName", "Set-CsOnlineApplicationInstance -Displayname `"$DisplayNameNormalised`"")) {
          Write-Verbose -Message "'$CurrentDisplayName' Changing DisplayName to: $DisplayNameNormalised"
          $null = (Set-CsOnlineApplicationInstance -Identity $UserPrincipalName -DisplayName "$DisplayNameNormalised" -ErrorAction STOP)
          Write-Information "SUCCESS: Displayname changed to '$DisplayName'"
          $CurrentDisplayName = $Object.DisplayName
        }
      }
      catch {
        Write-Verbose -Message 'FAILED - Error encountered changing DisplayName'
        Write-Error -Message 'Problem encountered with changing DisplayName' -Category NotImplemented -Exception $_.Exception -RecommendedAction 'Try manually with Set-CsOnlineApplicationInstance'
        Write-Debug $_
      }
    }
    #endregion

    #region Application Type
    if ($PSBoundParameters.ContainsKey('ApplicationType')) {
      $Operation = 'Application Type'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Application Type Change?
      if ($AppId -ne $CurrentAppId) {
        try {
          if ($PSCmdlet.ShouldProcess("$UserPrincipalName", "Set-CsOnlineApplicationInstance -ApplicationId $AppId")) {
            Write-Verbose -Message "'$Name ($UserPrincipalName )' Setting Application Type to: $ApplicationType"
            $null = (Set-CsOnlineApplicationInstance -Identity $UserPrincipalName -ApplicationId $AppId -ErrorAction STOP)
            Write-Verbose -Message 'SUCCESS'
          }
        }
        catch {
          Write-Error -Message 'Problem encountered changing Application Type' -Category NotImplemented -Exception $_.Exception -RecommendedAction 'Try manually with Set-CsOnlineApplicationInstance'
          Write-Debug $_
        }
      }
    }
    #endregion

    #region UsageLocation
    if ($PSBoundParameters.ContainsKey('UsageLocation')) {
      $Operation = 'Usage Location'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      if ($PSCmdlet.ShouldProcess("$UserPrincipalName", "Set-AzureADUser -UsageLocation $UsageLocation")) {
        try {
          Set-AzureADUser -ObjectId $UserPrincipalName -UsageLocation $UsageLocation -ErrorAction STOP
          Write-Verbose -Message "'$Name ($UserPrincipalName )' SUCCESS - Usage Location set to: $UsageLocation"
        }
        catch {
          if ($PSBoundParameters.ContainsKey('License')) {
            Write-Error -Message "'$Name ($UserPrincipalName )' Usage Location could not be set. Please apply manually before applying license" -Category NotSpecified -RecommendedAction 'Apply manually, then run Set-TeamsResourceAccount to apply license and phone number'
          }
          else {
            Write-Warning -Message "'$Name ($UserPrincipalName )' Usage Location cannot be set. If a license is needed, please assign UsageLocation manually beforehand"
          }
        }
      }
    }
    #endregion

    #region Licensing
    if ($PSBoundParameters.ContainsKey('License')) {
      # Verifying License is available to be assigned
      # Determining available Licenses from Tenant
      $Operation = 'Querying Licenses'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"
      $TenantLicenses = Get-TeamsTenantLicense

      # Changing License only if required
      $Operation = 'License Assignment'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      if ($License -eq $CurrentLicense) {
        # No action required
        Write-Information "'$Name ($UserPrincipalName )' License '$License' already assigned."
        $IsLicensed = $true
      }
      # Verifying License is available
      elseif ($License -eq 'PhoneSystemVirtualUser') {
        $RemainingPSVULicenses = ($TenantLicenses | Where-Object { $_.SkuPartNumber -eq 'PHONESYSTEM_VIRTUALUSER' }).Remaining
        Write-Verbose -Message "INFO: $RemainingPSVULicenses remaining Phone System Virtual User Licenses"
        if ($RemainingPSVULicenses -lt 1) {
          Write-Error -Message 'ERROR: No free PhoneSystem Virtual User License remaining in the Tenant.'
        }
        else {
          try {
            if ($PSCmdlet.ShouldProcess("$UserPrincipalName", 'Set-TeamsUserLicense -Add PhoneSystemVirtualUser')) {
              $null = (Set-TeamsUserLicense -Identity $UserPrincipalName -Add $License -ErrorAction STOP)
              Write-Information "'$Name ($UserPrincipalName )' License assignment - '$License' SUCCESS"
              $IsLicensed = $true
            }
          }
          catch {
            Write-Error -Message "'$Name ($UserPrincipalName )' License assignment failed for '$License'"
            Write-Debug $_
          }
        }
      }
      else {
        try {
          if ($PSCmdlet.ShouldProcess("$UPN", "Set-TeamsUserLicense -Add $License")) {
            $null = (Set-TeamsUserLicense -Identity $UPN -Add $License -ErrorAction STOP)
            Write-Information "'$Name ($UserPrincipalName )' License assignment - '$License' SUCCESS"
            $IsLicensed = $true
          }
        }
        catch {
          Write-Error -Message "'$Name ($UserPrincipalName )' License assignment failed for '$License'"
          Write-Debug $_
        }
      }
    }
    #endregion

    #region Waiting for License Application
    if ($PSBoundParameters.ContainsKey('License') -and $PSBoundParameters.ContainsKey('PhoneNumber')) {
      $Operation = 'Waiting for AzureAd to write Object'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      if ($License -eq 'PhoneSystemVirtualUser') {
        $ServicePlanName = 'MCOEV_VIRTUALUSER'
      }
      else {
        $ServicePlanName = 'MCOEV'
      }
      $i = 0
      $iMax = 600
      Write-Warning -Message "Applying a License may take longer than provisioned for ($($iMax/60) mins) in this Script - If so, please apply PhoneNumber manually with Set-TeamsResourceAccount"

      $Status = 'Applying License'
      $Operation = 'Waiting for Get-AzureAdUserLicenseDetail to return a Result'
      Write-Verbose -Message "$Status - $Operation"
      while (-not (Test-TeamsUserLicense -Identity $UserPrincipalName -ServicePlan $ServicePlanName)) {
        if ($i -gt $iMax) {
          Write-Error -Message "Could not find Successful Provisioning Status of the License '$ServicePlanName' in AzureAD in the last $iMax Seconds" -Category LimitsExceeded -RecommendedAction 'Please verify License has been applied correctly (Get-TeamsResourceAccount); Continue with Set-TeamsResourceAccount'
          return
        }
        Write-Progress -Id 1 -Activity 'Azure Active Directory is applying License. Please wait' `
          -Status $Status -SecondsRemaining $($iMax - $i) -CurrentOperation $Operation -PercentComplete (($i * 100) / $iMax)

        Start-Sleep -Milliseconds 1000
        $i++
      }
      Write-Progress -Id 1 -Activity 'Azure Active Directory is applying License. Please wait' -Status $Status -Completed
    }
    #endregion

    #region PhoneNumber
    if ($PSBoundParameters.ContainsKey('PhoneNumber')) {
      $Operation = 'Phone Number'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Removing old Number (if $null or different to current)
      #VALIDATE application of empty or Null does indeed remove the phonenumber!
      if ($force -or $null -eq $PhoneNumber -or $CurrentPhoneNumber -ne $PhoneNumber) {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' ACTION: Removing Phone Number"
        #CHECK PhoneNumber may not be removed correctly. Needs investigation
        try {
          if ($null -ne ($Object.TelephoneNumber)) {
            # Remove from VoiceApplicationInstance
            Write-Verbose -Message "'$Name ($UserPrincipalName )' Removing Microsoft Number"
            $null = (Set-CsOnlineVoiceApplicationInstance -Identity $UserPrincipalName -TelephoneNumber $null -WarningAction SilentlyContinue -ErrorAction STOP)
            Write-Verbose -Message 'SUCCESS'
          }
          if ($null -ne ($Object.OnPremLineURI)) {
            # Remove from ApplicationInstance
            Write-Verbose -Message "'$Name ($UserPrincipalName )' Removing Direct Routing Number"
            $null = (Set-CsOnlineApplicationInstance -Identity $UserPrincipalName -OnpremPhoneNumber $null -WarningAction SilentlyContinue -ErrorAction STOP)
            Write-Verbose -Message 'SUCCESS'
          }
        }
        catch {
          Write-Error -Message 'Removal of Number failed' -Category NotImplemented -Exception $_.Exception -RecommendedAction 'Try manually with Remove-AzureAdUser'
          Write-Debug $_
        }
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name)", (Get-CsOnlineApplicationInstance -Identity $UserPrincipalName | Select-Object UserPrincipalName, DisplayName, PhoneNumber | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }
      }
      else {
        Write-Verbose -Message "'$Name ($UserPrincipalName )' No Number assigned"
      }

      # Assigning Telephone Number
      if ($PhoneNumber) {
        if ($null -eq $CurrentLicense -and -not $IsLicensed) {
          Write-Error -Message 'A Phone Number can only be assigned to licensed objects.' -Category ResourceUnavailable -RecommendedAction 'Please apply a license before assigning the number. Set-TeamsResourceAccount can be used to do both'
        }
        else {
          Write-Verbose -Message "'$Name ($UserPrincipalName )' ACTION: Assigning Phone Number"
          # Assigning new Number
          # Processing paths for Telephone Numbers depending on Type
          try {
            # Loading all Microsoft Telephone Numbers
            if (-not $global:TeamsFunctionsMSTelephoneNumbers) {
              $global:TeamsFunctionsMSTelephoneNumbers = Get-CsOnlineTelephoneNumber -WarningAction SilentlyContinue
            }
            $MSNumber = ((Format-StringForUse -InputString "$PhoneNumber" -SpecialChars 'tel:+') -split ';')[0]

            if ($MSNumber -in $global:TeamsFunctionsMSTelephoneNumbers.Id) {
              # Set in VoiceApplicationInstance
              if ($PSCmdlet.ShouldProcess("$UserPrincipalName", "Set-CsOnlineVoiceApplicationInstance -Telephonenumber $E164Number")) {
                Write-Information "'$Name ($UserPrincipalName )' Number '$Number' found in Tenant, provisioning Microsoft for: Microsoft Calling Plans"
                $null = (Set-CsOnlineVoiceApplicationInstance -Identity $UserPrincipalName -TelephoneNumber $E164Number -ErrorAction STOP)
              }
            }
            else {
              # Set in ApplicationInstance
              if ($PSCmdlet.ShouldProcess("$UserPrincipalName", "Set-CsOnlineApplicationInstance -OnPremPhoneNumber $E164Number")) {
                Write-Information "'$Name ($UserPrincipalName )' Number '$E164Number' not found in Tenant, provisioning for: Direct Routing"
                $null = (Set-CsOnlineApplicationInstance -Identity $UserPrincipalName -OnpremPhoneNumber $E164Number -ErrorAction STOP)
              }
            }
          }
          catch {
            Write-Error -Message "'$Name ($UserPrincipalName )' Number '$PhoneNumber' not assigned!" -Category NotImplemented -RecommendedAction 'Please run Set-TeamsResourceAccount manually'
            Write-Debug $_
          }

        }
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name)", (Get-CsOnlineApplicationInstance -Identity $UserPrincipalName | Select-Object UserPrincipalName, DisplayName, PhoneNumber | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }
      }
    }
    #endregion
    #endregion

    Write-Progress -Id 1 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed

    if ( $PassThru ) {
      $Status = 'Output'
      $Operation = 'Querying Object'
      $step++
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      $RAObject = Get-TeamsResourceAccount -Identity $UserPrincipalName
      Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
      Write-Output $RAObject
    }
  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"

  } #end
} #Set-TeamsResourceAccount
