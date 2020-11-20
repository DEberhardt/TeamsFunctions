﻿# Module:   TeamsFunctions
# Function: ResourceAccount
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   BETA




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
  .LINK
    Get-TeamsResourceAccountAssociation
    New-TeamsResourceAccountAssociation
		Remove-TeamsResourceAccountAssociation
    New-TeamsResourceAccount
    Get-TeamsResourceAccount
    Set-TeamsResourceAccount
    Remove-TeamsResourceAccount
  #>
  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'CallQueue')]
  [Alias('New-TeamsRAAssoc')]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "UPN of the Object to change")]
    [string[]]$UserPrincipalName,

    [Parameter(Mandatory = $true, ParameterSetName = 'CallQueue', Position = 1, ValueFromPipelineByPropertyName = $true, HelpMessage = "Name of the CallQueue")]
    [string]$CallQueue,

    [Parameter(Mandatory = $true, ParameterSetName = 'AutoAttendant', Position = 1, ValueFromPipelineByPropertyName = $true, HelpMessage = "Name of the AutoAttendant")]
    [string]$AutoAttendant,

    [Parameter(Mandatory = $false)]
    [switch]$Force
  ) #param

  begin {
    # Caveat - Script in Development
    $VerbosePreference = "Continue"
    $DebugPreference = "Continue"
    Show-FunctionStatus -Level BETA
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

    # Enabling $Confirm to work with $Force
    if ($Force -and -not $Confirm) {
      $ConfirmPreference = 'None'
    }

    # Initialising counters for Progress bars - Level 0
    [int]$step = 0
    [int]$sMax = 4

    #region Determining and Validating Entity
    # Determining $EntityObject
    $Operation = "Determining Entity"
    Write-Progress -Id 0 -Status $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    switch ($PSCmdlet.ParameterSetName) {
      'CallQueue' {
        $DesiredType = "CallQueue"
        $Entity = $CallQueue
        # Querying Call Queue by Name - need Unique Result
        Write-Verbose -Message "Querying Call Queue '$CallQueue'"
        $EntityObject = Get-CsCallQueue -NameFilter "$CallQueue" -WarningAction SilentlyContinue
      }
      'AutoAttendant' {
        $DesiredType = "AutoAttendant"
        $Entity = $AutoAttendant
        # Querying Auto Attendant by Name - need Unique Result
        Write-Verbose -Message "Querying Auto Attendant '$AutoAttendant'"
        $EntityObject = Get-CsAutoAttendant -NameFilter "$AutoAttendant" -WarningAction SilentlyContinue
      }
    }

    # Validating Unique result received
    $Operation = "Validating whether a Unique result is obtained"
    $step++
    Write-Progress -Id 0 -Status $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    if ($null -eq $EntityObject) {
      Write-Error "$DesiredType '$Entity' - Not found" -Category ParserError -RecommendedAction "Please check 'DesiredType' exists with this Name"
      return
    }
    elseif ($EntityObject.GetType().BaseType.Name -eq "Array") {
      Write-Verbose -Message "'$Entity' - Multiple results found! This script is based on lookup via Name, which requires, for safety reasons, a unique Name to process." -Verbose
      Write-Verbose -Message "Here are all objects found with the Name. Please use the correct Identity to run New-CsOnlineApplicationInstanceAssociation!" -Verbose
      $EntityObject | Select-Object Identity, Name | Format-Table
      Write-Error "$DesiredType '$Entity' - Multiple Results found! Cannot determine unique result. Please use New-CsOnlineApplicationInstanceAssociation!" -Category ParserError -RecommendedAction "Please use New-CsOnlineApplicationInstanceAssociation!"
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
    $Operation = "Processing provided UserPrincipalNames"
    $step++
    Write-Progress -Id 0 -Status $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
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

    $Operation = "Processing found Resource Accounts"
    $step++
    Write-Progress -Id 0 -Status $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $Counter = 0
    foreach ($Account in $Accounts) {
      Write-Progress -Id 1 -Status "Validating '$($Account.UserPrincipalName)'" -Activity $MyInvocation.MyCommand -PercentComplete ($Counter / $($Accounts.Count) * 100)
      $Counter++
      # Query existing connection

      # Initialising counters for Progress bars - Level 0
      [int]$step2 = 1
      [int]$sMax2 = 5

      $Operation = "Querying existing associations"
      Write-Progress -Id 2 -Status "Validating '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
      Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
      #CHECK Query rather with Get-CsOnlineApplicationInstanceAssociation? - Needs ObjectId though! (replicated from Get-TeamsResourceAccountAssociation)
      #TODO Test Performance of GET-TeamsResourceAccountAssociation VS Get-CsOnlineApplicationInstanceAssociation
      <# Needs testing
      $ExistingConnection = Get-CsOnlineApplicationInstanceAssociation (Get-AzureAdUser -ObjectId $Account.UserPrincipalName) -WarningAction SilentlyContinue
      #>
      $ExistingConnection = $null
      $ExistingConnection = Get-TeamsResourceAccountAssociation $Account.UserPrincipalName -WarningAction SilentlyContinue
      if ($null -eq $ExistingConnection.AssociatedTo) {
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - No assignment found. OK"
      }
      else {
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - This account is already assigned to the following entity:" -Verbose
        $ExistingConnection
        Write-Error -Message "'$($Account.UserPrincipalName)' - This account cannot be associated as it is already assigned to $($ExistingConnection.ConfigurationType) '$($ExistingConnection.AssociatedTo)'"
        #TEST this
        [void]$Accounts.Remove($Account)
      }

      # Comparing ApplicationType
      $Operation = "Validating ApplicationType"
      $step2++
      Write-Progress -Id 2 -Status "Validating '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
      Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
      #TEST this. Was restructured from below. needs testing.
      if ((Get-TeamsResourceAccount -Identity $Account.UserPrincipalName -WarningAction SilentlyContinue).ApplicationType -eq $DesiredType) {
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - Application type matches '$DesiredType' - OK"
      }
      else {
        if ($PSBoundParameters.ContainsKey('Force')) {
          # Changing Application Type
          $Operation = "Application Type is not '$DesiredType' - Changing"
          $step2++
          Write-Progress -Id 2 -Status "Validating '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
          try {
            $null = Set-CsOnlineApplicationInstance -Identity $Account.ObjectId -ApplicationId $(GetAppIdFromApplicationType $DesiredType) -ErrorAction Stop
          }
          catch {
            Write-Error -Message "'$($Account.UserPrincipalName)' - Application type does not match and could not be changed! Expected: '$DesiredType' - Please change manually or recreate the Account" -Category InvalidType -RecommendedAction "Please change manually or recreate the Account"
            continue
          }

          $Operation = "Application Type is not '$DesiredType' - Waiting for AzureAD (2s)"
          $step2++
          Write-Progress -Id 2 -Status "Validating '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
          Start-Sleep -Seconds 2

          $Operation = "Application Type is not '$DesiredType' - Verifying"
          $step2++
          Write-Progress -Id 2 -Status "Validating '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
          Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
          if ($DesiredType -ne $(GetApplicationTypeFromAppId (Get-CsOnlineApplicationInstance -Identity $Account.ObjectId -WarningAction SilentlyContinue).ApplicationId)) {
            Write-Error -Message "'$($Account.UserPrincipalName)' - Application type could not be changed to Desired Type: '$DesiredType'" -Category InvalidType -ErrorAction Stop
          }
          else {
            Write-Verbose -Message "'$($Account.UserPrincipalName)' - Changing Application Type to '$DesiredType': SUCCESS" -Verbose
          }
        }
        else {
          Write-Error -Message "'$($Account.UserPrincipalName)' - Application type does not match! Expected: '$DesiredType' - Please change manually or use -Force switch" -Category InvalidType -RecommendedAction "Please change manually or use -Force switch"
          continue
        }
      }
      Write-Progress -Id 2 -Status "'$($Account.UserPrincipalName)' - Complete" -Activity $MyInvocation.MyCommand -Completed
    }
    Write-Progress -Id 1 -Status "Validation Complete" -Activity $MyInvocation.MyCommand -Completed

    # Processing found accounts
    if ( $Accounts ) {
      # Processing Assignment
      Write-Verbose -Message "Processing assignment of all Accounts to $DesiredType '$($EntityObject.Name)'"
      $Counter = 0
      foreach ($Account in $Accounts) {
        Write-Progress -Id 1 -Status "Processing assignment of '$($Account.UserPrincipalName)'" -Activity $MyInvocation.MyCommand -PercentComplete ($Counter / $($Accounts.Count) * 100)
        $Counter++

        # Initialising counters for Progress bars - Level 0
        [int]$step3 = 1
        [int]$sMax3 = 4

        # Establishing Association
        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)'"
        Write-Progress -Id 2 -Status "Processing assignment of '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step3 / $sMax3 * 100)
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"

        if ($PSCmdlet.ShouldProcess("$($Account.UserPrincipalName)", "New-CsOnlineApplicationInstanceAssociation")) {
          $OperationStatus = New-CsOnlineApplicationInstanceAssociation -Identities $Account.ObjectId -ConfigurationType $DesiredType -ConfigurationId $EntityObject.Identity
        }

        # Re-query Association Target
        #  Wating for AAD to write the Association Target so that it may be queried correctly
        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)' - Waiting for AzureAD (2s)"
        $step3++
        Write-Progress -Id 2 -Status "Processing assignment of '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
        Start-Sleep -Seconds 2


        $Operation = "Assigning to $DesiredType '$($EntityObject.Name)' - Verifying"
        $step3++
        Write-Progress -Id 2 -Status "Processing assignment of '$($Account.UserPrincipalName)'" -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step2 / $sMax2 * 100)
        Write-Verbose -Message "'$($Account.UserPrincipalName)' - $Operation"
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

        Write-Progress -Id 2 -Status "'$($Account.UserPrincipalName)' - Complete" -Activity $MyInvocation.MyCommand -Completed

        Write-Output $ResourceAccountAssociationObject
      }
      Write-Progress -Id 1 -Status "'$($Account.UserPrincipalName)' - Complete" -Activity $MyInvocation.MyCommand -Completed
    }
    else {
      Write-Warning -Message "No Accounts found"
    }

    Write-Progress -Id 0 -Status "Complete" -Activity $MyInvocation.MyCommand -Completed

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #New-TeamsResourceAccountAssociation
