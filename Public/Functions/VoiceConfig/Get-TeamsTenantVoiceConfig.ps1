﻿# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:		David Eberhardt
# Updated:  01-OCT-2020
# Status:   PreLive




function Get-TeamsTenantVoiceConfig {
  <#
	.SYNOPSIS
		Displays Information about available Voice Configuration in the Tenant
	.DESCRIPTION
		Displays all Voice relevant information configured in the Tenant incl. counters for free Licenses and Numbers
  .PARAMETER DisplayUserCounters
    Optional. Displays information about Users enabled for Teams and for EnterpriseVoice
    This extends Script execution depending on number of Users in the Tenant
  .PARAMETER Detailed
    Optional. Displays more information about Voice Routing Policies, Dial Plans, etc.
	.EXAMPLE
		Get-TeamsTenantVoiceConfig
    Displays Licenses for Call Plans, available Numbers, as well as
    Counters for all relevant Policies, available VoiceRoutingPolicies
	.EXAMPLE
		Get-TeamsTenantVoiceConfig -Detailed
    Displays a detailed view also listing Names for DialPlans, PSTN Usages, Voice Routes and PSTN Gateways
    Also displays diagnostic parameters for troubleshooting
	.NOTES
		General notes
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Find-TeamsUserVoiceConfig
  .LINK
    Get-TeamsTenantVoiceConfig
  .LINK
    Get-TeamsUserVoiceConfig
  .LINK
    Set-TeamsUserVoiceConfig
  .LINK
    Remove-TeamsUserVoiceConfig
  .LINK
    Test-TeamsUserVoiceConfig
  #>

  [CmdletBinding()]
  param(
    [Parameter(HelpMessage = 'Displays counters for User information')]
    [switch]$DisplayUserCounters,

    [Parameter(HelpMessage = 'Displays detailed information')]
    [switch]$Detailed
  ) #param

  begin {
    Show-FunctionStatus -Level PreLive
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Asserting SkypeOnline Connection
    if (-not (Assert-SkypeOnlineConnection)) { break }

    # Initialising counters for Progress bars
    [int]$step = 0
    [int]$sMax = 4
    if ( $DisplayUserCounters ) { $sMax = $sMax + 3 }
    if ( $Detailed ) { $sMax++ }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    #region Information Gathering
    $Operation = 'Querying Tenant'
    Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $Tenant = Get-CsTenant -WarningAction SilentlyContinue

    $Operation = 'Querying SIP Domains'
    $step++
    Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $SipDomains = Get-CsOnlineSipDomain -WarningAction SilentlyContinue

    $Operation = 'Querying Tenant Licenses'
    $step++
    Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $TenantLicenses = Get-TeamsTenantLicense
    $CallPlanINT = $TenantLicenses | Where-Object SkuPartNumber -EQ 'MCOPSTN1'
    $CallPlanDOM = $TenantLicenses | Where-Object SkuPartNumber -EQ 'MCOPSTN2'
    $CallPlanDOM120 = $TenantLicenses | Where-Object { $_.SkuPartNumber -EQ 'MCOPSTN5' -or $_.SkuPartNumber -EQ 'MCOPSTN_5' }
    $CommunicationC = $TenantLicenses | Where-Object SkuPartNumber -EQ 'MCOPSTNC'

    $Operation = 'Querying Direct Routing Information'
    $step++
    Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $TDP = Get-CsTenantDialPlan -WarningAction SilentlyContinue
    $OVP = Get-CsOnlineVoiceRoutingPolicy -WarningAction SilentlyContinue
    $OPU = (Get-CsOnlinePSTNusage -WarningAction SilentlyContinue).Usage
    $OVR = Get-CsOnlineVoiceRoute -WarningAction SilentlyContinue
    $OGW = Get-CsOnlinePSTNGateway -WarningAction SilentlyContinue
    #endregion

    #region Creating Base Custom Object
    $Operation = 'Building Base Object'
    $step++
    Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message $Operation
    $Object = [PSCustomObject][ordered]@{
      DisplayName                            = $Tenant.DisplayName
      Domains                                = $Tenant.Domains
      SipDomains                             = $SipDomains.Name
      TeamsUpgradeEffectiveMode              = $Tenant.TeamsUpgradeEffectiveMode
      TenantLicenses                         = $TenantLicenses.ProductName
      InternationalCallingPlanUnitsRemaining = $CallPlanINT.Remaining
      DomesticCallingPlanUnitsRemaining      = $CallPlanDOM.Remaining
      DomesticCallingPlan120UnitsRemaining   = $CallPlanDOM120.Remaining
      CommunicationCreditsUnitsRemaining     = $CommunicationC.Remaining
      ConfiguredTenantDialPlans              = $TDP.Count
      ConfiguredOnlineVoiceRoutingPolicies   = $OVP.Count
      ConfiguredOnlinePSTNUsages             = $OPU.Count
      ConfiguredOnlineVoiceRoutes            = $OVR.Count
      ConfiguredOnlinePSTNGateways           = $OGW.Count
    }
    #endregion

    #region User Information
    if ($PSBoundParameters.ContainsKey('DisplayUserCounters')) {
      Write-Verbose -Message 'DisplayUserCounters - Querying User Information - This will take some time!' -Verbose

      $Operation = 'DisplayUserCounters - Querying AzureADUsers'
      $step++
      Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message $Operation
      $AdUsers = Get-AzureADUser -All:$TRUE | Where-Object AccountEnabled -EQ $TRUE -WarningAction SilentlyContinue

      $Operation = 'DisplayUserCounters - Querying CsOnlineUsers'
      $step++
      Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message $Operation
      $CsOnlineUsers = Get-CsOnlineUser -WarningAction SilentlyContinue

      $Operation = 'DisplayUserCounters - Counting EV Users'
      $step++
      Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message $Operation
      $CsOnlineUsersEV = $CsOnlineUsers | Where-Object EnterpriseVoiceEnabled -EQ $TRUE

      $Object | Add-Member -MemberType NoteProperty -Name UsersEnabledInAzureAD -Value $AdUsers.Count
      $Object | Add-Member -MemberType NoteProperty -Name UsersEnabledForTeams -Value $CsOnlineUsers.Count
      $Object | Add-Member -MemberType NoteProperty -Name UsersEnabledForEnterpriseVoice -Value $CsOnlineUsersEV.Count

    }
    #endregion

    #region Detailed Information
    if ($PSBoundParameters.ContainsKey('Detailed')) {
      $Operation = 'Detailed - Querying Microsoft Telephone Numbers Information'
      $step++
      Write-Progress -Id 0 -Status 'Information Gathering' -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message $Operation
      if (-not $global:TeamsFunctionsMSTelephoneNumbers) {
        $global:TeamsFunctionsMSTelephoneNumbers = Get-CsOnlineTelephoneNumber -WarningAction SilentlyContinue
      }

      if ( $null -ne $global:TeamsFunctionsMSTelephoneNumbers ) {
        $MSTelephoneNumbersCount = $global:TeamsFunctionsMSTelephoneNumbers.Count
        [int]$MSTelephoneNumbersFree = ($global:TeamsFunctionsMSTelephoneNumbers | Where-Object TargetType -NE $null).Count

        $MSNumbersUser = $global:TeamsFunctionsMSTelephoneNumbers | Where-Object InventoryType -EQ 'Subscriber'
        [int]$MSTelephoneNumbersUser = $MSNumbersUser.Count
        [int]$MSTelephoneNumbersUserFree = ($MSNumbersUser | Where-Object TargetType -NE $null).Count

        $MSNumbersService = $global:TeamsFunctionsMSTelephoneNumbers | Where-Object InventoryType -EQ 'Service'
        [int]$MSTelephoneNumbersService = $MSNumbersService.Count
        [int]$MSTelephoneNumbersServiceFree = ($MSNumbersService | Where-Object TargetType -NE $null).Count

        $MSNumbersTollFree = $global:TeamsFunctionsMSTelephoneNumbers | Where-Object InventoryType -EQ 'TollFree'
        [int]$MSTelephoneNumbersTollFree = $MSNumbersTollFree.Count
        [int]$MSTelephoneNumbersTollFreeFree = ($MSNumbersTollFree | Where-Object TargetType -NE $null).Count

      }
      else {
        $MSTelephoneNumbersCount = 0
        $MSTelephoneNumbersFree = 0
        $MSTelephoneNumbersUser = 0
        $MSTelephoneNumbersUserFree = 0
        $MSTelephoneNumbersService = 0
        $MSTelephoneNumbersServiceFree = 0
        $MSTelephoneNumbersTollFree = 0
        $MSTelephoneNumbersTollFreeFree = 0
      }

      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbers -Value $MSTelephoneNumbersCount
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersFree -Value $MSTelephoneNumbersFree
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersUser -Value $MSTelephoneNumbersUser
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersUserFree -Value $MSTelephoneNumbersUserFree
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersService -Value $MSTelephoneNumbersService
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersServiceFree -Value $MSTelephoneNumbersServiceFree
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersTollFree -Value $MSTelephoneNumbersTollFree
      $Object | Add-Member -MemberType NoteProperty -Name MSTelephoneNumbersTollFreeFree -Value $MSTelephoneNumbersTollFreeFree
      $Object | Add-Member -MemberType NoteProperty -Name TenantDialPlans -Value $TDP.Identity
      $Object | Add-Member -MemberType NoteProperty -Name OnlineVoiceRoutingPolicies -Value $OVP.Identity
      $Object | Add-Member -MemberType NoteProperty -Name OnlinePSTNUsages -Value $OPU
      $Object | Add-Member -MemberType NoteProperty -Name OnlineVoiceRoutes -Value $OVR.Identity
      $Object | Add-Member -MemberType NoteProperty -Name OnlinePSTNGateways -Value $OGW.Identity
      $Object | Add-Member -MemberType NoteProperty -Name DirSyncEnabled -Value $Tenant.DirSyncEnabled
      $Object | Add-Member -MemberType NoteProperty -Name LastSyncTimeStamp -Value $Tenant.LastSyncTimeStamp

    }
    #endregion

    # Output
    Write-Progress -Id 0 -Status 'Information Gathering' -Activity $MyInvocation.MyCommand -Completed
    Write-Output $Object

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #Get-TeamsTenantVoiceConfig
