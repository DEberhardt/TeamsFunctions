# Module:   TeamsFunctions
# Function: ResourceAccount
# Author:		David Eberhardt
# Updated:  01-DEC-2020
# Status:   Live




function New-TeamsResourceAccountAssociation {
  <#
	.SYNOPSIS
		Connects one or more Resource Accounts to a single CallQueue or AutoAttendant
	.DESCRIPTION
		Associates one or more existing Resource Accounts to a Call Queue or Auto Attendant
		Resource Account Type is checked against the ApplicationType.
		User is prompted if types do not match
	.PARAMETER UserPrincipalName
		Required. UPN(s) of the Resource Account(s) to be associated to a Call Queue or AutoAttendant
	.PARAMETER CallQueue
		Optional. Specifies the connection to be made to the provided Call Queue Name
	.PARAMETER AutoAttendant
		Optional. Specifies the connection to be made to the provided Auto Attendant Name
	.PARAMETER Force
		Optional. Suppresses Confirmation dialog if -Confirm is not provided
		Used to override prompts for alignment of ApplicationTypes.
		The Resource Account is changed to have the same type as the associated Object (CallQueue or AutoAttendant)!
	.EXAMPLE
		New-TeamsResourceAccountAssociation -UserPrincipalName Account1@domain.com -
		Explanation of what the example does
  .INPUTS
    System.String
  .OUTPUTS
    System.Object
	.NOTES
    Connects multiple Resource Accounts to ONE CallQueue or AutoAttendant
    The Type of the Resource Account has to corellate to the entity connected.
    Parameter Force can be used to change the type of RA to align to the entity if possible.
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
    Set-TeamsResourceAccount
	.LINK
    Remove-TeamsResourceAccount
  #>
  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'CallQueue')]
  [Alias('New-TeamsRAA')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'UPN of the Object to change')]
    [string[]]$UserPrincipalName,

    [Parameter(Mandatory, ParameterSetName = 'CallQueue', ValueFromPipelineByPropertyName, HelpMessage = 'Name of the CallQueue')]
    [string]$CallQueue,

    [Parameter(Mandatory, ParameterSetName = 'AutoAttendant', ValueFromPipelineByPropertyName, HelpMessage = 'Name of the AutoAttendant')]
    [string]$AutoAttendant,

    [Parameter(Mandatory = $false)]
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

    # Enabling $Confirm to work with $Force
    if ($Force -and -not $Confirm) {
      $ConfirmPreference = 'None'
    }

    # Initialising counters for Progress bars - Level 0
    [int]$step = 0
    [int]$sMax = 5

    #region Determining and Validating Entity
    # Determining $EntityObject
    $Status = 'Validation'
    $Operation = 'Determining Entity'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    try {
      switch ($PSCmdlet.ParameterSetName) {
        'CallQueue' {
          $DesiredType = 'CallQueue'
          $Entity = $CallQueue
          # Querying Call Queue by Name - need Unique Result
          Write-Verbose -Message "Querying Call Queue '$CallQueue'"
          $EntitySearch = Get-CsCallQueue -NameFilter "$CallQueue" -WarningAction SilentlyContinue
        }
        'AutoAttendant' {
          $DesiredType = 'AutoAttendant'
          $Entity = $AutoAttendant
          # Querying Auto Attendant by Name - need Unique Result
          Write-Verbose -Message "Querying Auto Attendant '$AutoAttendant'"
          $EntitySearch = Get-CsAutoAttendant -NameFilter "$AutoAttendant" -WarningAction SilentlyContinue
        }
      }
    }
    catch {
      #TEST this should address the issue with Script not stopping if Entity not found.
      throw "Cannot determine $DesiredType '$Entity'"
      #Write-Error -Message "Cannot determine $DesiredType '$Entity'" -ErrorAction Stop
    }
    if ($EntitySearch.Count -gt 1) {
      #NOTE Displayname is not normalised - can lead to inconsistencies - it shouldn't need to, so needs a warning - if not found, try normalising!
      $EntityObject = $EntitySearch | Where-Object Name -EQ "$Entity"
    }
    else {
      $EntityObject = $EntitySearch
    }

    # Validating Unique result received
    $Operation = 'Unique result'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    if ($null -eq $EntityObject) {
      #throw "$DesiredType '$Entity' - Not found, please check entity exists with this Name"
      Write-Error -Exception 'ObjectNotFound' -Message "$DesiredType '$Entity' - Not found, please check entity exists with this Name"
      return
    }
    elseif ($EntityObject -is [Array]) {
      $EntityObject = $EntityObject | Where-Object Name -EQ "$Entity"
      Write-Verbose -Message "'$Entity' - Multiple results found! This script is based on lookup via Name, which requires, for safety reasons, a unique Name to process." -Verbose
      Write-Verbose -Message 'Listing all objects found with the Name. Please use the correct Identity to run New-CsOnlineApplicationInstanceAssociation!' -Verbose
      $EntityObject | Select-Object Identity, Name | Format-Table
      Write-Error "$DesiredType '$Entity' - Multiple Results found! Cannot determine unique result. Please use New-CsOnlineApplicationInstanceAssociation!" -Category ParserError -RecommendedAction 'Please use New-CsOnlineApplicationInstanceAssociation!'
      return
    }
    else {
      Write-Verbose -Message "$DesiredType '$Entity' - Unique result found: $($EntityObject.Name)"
    }
    #endregion

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    # Query $UserPrincipalName
    [System.Collections.ArrayList]$Accounts = @()
    $Operation = 'Processing provided UserPrincipalNames'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    #TEST If AA or RA are not found, should not throw multiple errors, but skip this Entity or RA
    foreach ($UPN in $UserPrincipalName) {
      Write-Verbose -Message "Querying Resource Account '$UPN'"
      try {
        $RAObject = Get-AzureADUser -ObjectId $UPN -WarningAction SilentlyContinue -ErrorAction Stop
        $AppInstance = Get-CsOnlineApplicationInstance $RAObject.ObjectId -WarningAction SilentlyContinue -ErrorAction Stop
        [void]$Accounts.Add($AppInstance)
        Write-Verbose "Resource Account found: '$($AppInstance.DisplayName)'"
      }
      catch {
        Write-Error "Resource Account not found: '$UPN'" -Category ObjectNotFound
        continue
      }
    }

    # Breaks the chain if no eligible accounts are found.
    if ( -not $Accounts ) {
      Write-Warning -Message 'No Resource Accounts found eligible for Association. Stopping.'
      return
    }

    $Operation = 'Processing found Resource Accounts'
    $step++
    Write-Progress -Id 0 -Status $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $Counter = 1
    [System.Collections.ArrayList]$ValidatedAccounts = @()
    foreach ($Account in $Accounts) {
      $Status = 'Processing'
      $Operation = "'$($Account.UserPrincipalName)'"
      Write-Progress -Id 0 -Status $Status -Activity $MyInvocation.MyCommand -PercentComplete ($Counter / $($Accounts.Count) * 100)
      Write-Verbose -Message "$Status - $Operation"
      $Counter++
      # Query existing connection

      # Initialising counters for Progress bars - Level 0
      [int]$step2 = 1
      [int]$sMax2 = 6

      $Status = "'$($Account.UserPrincipalName)'"
      $Operation = 'Querying existing associations'
      Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
      Write-Verbose -Message "$Status - $Operation"
      $ExistingConnection = $null
      $ExistingConnection = Get-CsOnlineApplicationInstanceAssociation -Identity $Account.ObjectId -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

      if ($null -eq $ExistingConnection.ConfigurationId) {
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - No assignment found. OK"
      }
      else {
        Write-Error -Message "'$($Account.UserPrincipalName)' - This account cannot be associated as it is already assigned as '$($ExistingConnection.ConfigurationType)'"
        continue
      }

      # Comparing ApplicationType
      $Operation = 'Validating ApplicationType'
      $step2++
      Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
      Write-Verbose -Message "$Status - $Operation"
      $ApplicationTypeMatches = ((Get-CsOnlineApplicationInstance -Identity $Account.UserPrincipalName -WarningAction SilentlyContinue).ApplicationId -eq (GetAppIdFromApplicationType $DesiredType))

      if ( $ApplicationTypeMatches ) {
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - Application type matches '$DesiredType' - OK"
      }
      else {
        if ($PSBoundParameters.ContainsKey('Force')) {
          # Changing Application Type
          $Operation = "Application Type is not '$DesiredType' - Changing"
          $step2++
          Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "$Status - $Operation"
          try {
            $null = Set-CsOnlineApplicationInstance -Identity $Account.ObjectId -ApplicationId $(GetAppIdFromApplicationType $DesiredType) -ErrorAction Stop
          }
          catch {
            Write-Error -Message "'$($Account.UserPrincipalName)' - Application type does not match and could not be changed! Expected: '$DesiredType' - Please change manually or recreate the Account" -Category InvalidType -RecommendedAction 'Please change manually or recreate the Account'
            continue
          }

          $Operation = "Application Type is not '$DesiredType' - Waiting for AzureAD (2s)"
          $step2++
          Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "$Status - $Operation"
          Start-Sleep -Seconds 2

          $Operation = "Application Type is not '$DesiredType' - Verifying"
          $step2++
          Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "$Status - $Operation"
          if ($DesiredType -ne $(GetApplicationTypeFromAppId (Get-CsOnlineApplicationInstance -Identity $Account.ObjectId -WarningAction SilentlyContinue).ApplicationId)) {
            Write-Error -Message "'$($Account.UserPrincipalName)' - Application type could not be changed to Desired Type: '$DesiredType'" -Category InvalidType
            continue
          }
          else {
            Write-Information "'$($Account.UserPrincipalName)' - Changing Application Type to '$DesiredType': SUCCESS"
          }
        }
        else {
          Write-Warning -Message "'$($Account.UserPrincipalName)' - Application type does not match! Expected '$DesiredType' - Omitting account. Please change type manually or use -Force switch"
          continue
        }
      }

      [void]$ValidatedAccounts.Add($Account)
      Write-Progress -Id 1 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    }

    # Processing found accounts
    if ( $ValidatedAccounts ) {
      # Processing Assignment
      Write-Verbose -Message "Processing assignment of all Accounts to $DesiredType '$($EntityObject.Name)'"
      $Counter = 1
      foreach ($Account in $Accounts) {
        $ErrorEncountered = $null

        $Status = 'Assignment'
        $Operation = "'$($Account.UserPrincipalName)'"
        Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($Counter / $($Accounts.Count) * 100)
        Write-Verbose -Message "$Status - $Operation"
        $Counter++

        # Initialising counters for Progress bars - Level 0
        [int]$step3 = 1
        [int]$sMax3 = 4

        # Establishing Association
        $Status = "'$($Account.UserPrincipalName)'"
        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)'"
        Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step3 / $sMax3 * 100)
        Write-Verbose -Message "$Status - $Operation"

        # Creating Splatting Object
        $Parameters = $null
        $Parameters += @{ 'Identities' = $Account.ObjectId }
        $Parameters += @{ 'ConfigurationType' = $DesiredType }
        $Parameters += @{ 'ConfigurationId' = $EntityObject.Identity }
        $Parameters += @{ 'ErrorAction' = 'Stop' }

        # Create CsAutoAttendantCallableEntity
        Write-Verbose -Message '[PROCESS] Creating Resource Account Association'
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name): Parameters:", ($Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }

        if ($PSCmdlet.ShouldProcess("$($Account.UserPrincipalName)", 'New-CsOnlineApplicationInstanceAssociation')) {
          #$OperationStatus = New-CsOnlineApplicationInstanceAssociation -Identities $Account.ObjectId -ConfigurationType $DesiredType -ConfigurationId $EntityObject.Identity
          try {
            $OperationStatus = New-CsOnlineApplicationInstanceAssociation @Parameters
          }
          catch {
            $ErrorEncountered = $_
          }
        }

        # Re-query Association Target
        #  Wating for AAD to write the Association Target so that it may be queried correctly
        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)' - Waiting for AzureAD (2s)"
        $step3++
        Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step3 / $sMax3 * 100)
        Write-Verbose -Message "$Status - $Operation"
        Start-Sleep -Seconds 2

        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)' - Verifying"
        $step3++
        Write-Progress -Id 1 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step3 / $sMax3 * 100)
        Write-Verbose -Message "$Status - $Operation"
        $AssociationTarget = switch ($PSCmdlet.ParameterSetName) {
          'CallQueue' {
            Get-CsCallQueue -Identity $OperationStatus.Results.ConfigurationId -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
          }
          'AutoAttendant' {
            Get-CsAutoAttendant -Identity $OperationStatus.Results.ConfigurationId -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
          }
        }

        # Output
        $ResourceAccountAssociationObject = $null
        $ResourceAccountAssociationObject = [PSCustomObject][ordered]@{
          UserPrincipalName = $Account.UserPrincipalName
          ConfigurationType = $OperationStatus.Results.ConfigurationType
          Result            = $OperationStatus.Results.Result
          StatusCode        = $OperationStatus.Results.StatusCode
          StatusMessage     = $OperationStatus.Results.Message
          AssociatedTo      = $AssociationTarget.Name
        }

        Write-Progress -Id 1 -Status "'$($Account.UserPrincipalName)' - Complete" -Activity $MyInvocation.MyCommand -Completed
        Write-Output $ResourceAccountAssociationObject

        if ( $ErrorEncountered ) {
          Write-Error -Message "Association of Object failed with exception: $($ErrorEncountered.Exception.Message)" -ErrorAction Stop
        }
      }
    }

  } #process

  end {
    Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsResourceAccountAssociation
