﻿# Module:   TeamsFunctions
# Function: Licensing
# Author:		Philipp, Scripting.up-in-the.cloud
# Updated:  01-DEC-2020
# Status:   PreLive




function Get-AzureAdLicense {
  <#
	.SYNOPSIS
    License information for AzureAD Licenses related to Teams
  .DESCRIPTION
    Returns an Object containing all Teams related Licenses
  .PARAMETER FilterRelevantForTeams
    Optional. By default, shows all 365 Licenses
    Using this switch, shows only Licenses relevant for Teams
  .EXAMPLE
    Get-AzureAdLicense
    Returns 39 Azure AD Licenses that relate to Teams for use in other commands
  .NOTES
    Source
    https://scripting.up-in-the.cloud/licensing/o365-license-names-its-a-mess.html
    With very special thanks to Philip
    Reads
    https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
  .COMPONENT
    Teams Migration and Enablement. License Assignment
  .ROLE
    Licensing
  .FUNCTIONALITY
		Returns a list of Licenses
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Get-TeamsTenantLicense
  .LINK
    Get-TeamsUserLicense
  .LINK
    Set-TeamsUserLicense
  .LINK
    Test-TeamsUserLicense
  .LINK
    Get-AzureAdLicense
  .LINK
    Get-AzureAdLicenseServicePlan
  #>

  [CmdletBinding()]
  [OutputType([Object[]])]
  param(
    [Parameter()]
    [switch]$FilterRelevantForTeams
  ) #param

  begin {
    Show-FunctionStatus -Level PreLive
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

    [System.Collections.ArrayList]$Products = @()

    $srcProductPlans = @{}
    $planServicePlanNames = @{}

    [System.Collections.ArrayList]$ProductsNotAdded = @()
    [System.Collections.ArrayList]$PlansNotAdded = @()

  } #begin

  process {
    #read the content of the Microsoft web page and extract the first table
    $url = 'https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $content = (Invoke-WebRequest $url -UseBasicParsing).Content
    $content = $content.SubString($content.IndexOf('<tbody>'))
    $content = $content.Substring(0, $content.IndexOf('</tbody>'))

    #eliminate line feeds so that we can use regular expression to get the table rows...
    $content = $content -replace "`r?`n", ''
    $rows = (Select-String -InputObject $content -Pattern '<tr>(.*?)</tr>' -AllMatches).Matches | ForEach-Object {
      $_.Groups[1].Value
    }

    #on each table row, get the column cell content
    #   1st cell contains the product display name
    #   2nd cell contains the Sku ID (called 'string ID' here)
    #   3rd cell contains the included service plans (with string IDs)
    #   3rd cell contains the included service plans (with display names)
    $rows | ForEach-Object {
      $cells = (Select-String -InputObject $_ -Pattern '<td>(.*?)</td>' -AllMatches).Matches | ForEach-Object {
        $_.Groups[1].Value
      }

      $srcProductName = $cells[0]
      $srcSkuPartNumber = $cells[1]
      $srcSkuId = $cells[2]
      $srcServicePlan = $cells[3]
      $srcServicePlanName = $cells[4]

      $srcProductPlans = $null
      [System.Collections.ArrayList]$srcProductPlans = @()

      #region Sub-Skus (Plans)
      # Preparing Plans
      if (($srcServicePlan.Trim() -ne '') -and ($srcServicePlanName.Trim() -ne '')) {

        #store the service plan string IDs for later match
        $srcServicePlan -split '<br.?>' | ForEach-Object {
          $planServicePlanName = ($_.SubString(0, $_.LastIndexOf('('))).Trim()
          $planServicePlanId = $_.SubString($_.LastIndexOf('(') + 1)
          if ($planServicePlanId.Contains(')')) {
            $planServicePlanId = $planServicePlanId.SubString(0, $planServicePlanId.IndexOf(')'))
          }

          if (-not $planServicePlanNames.ContainsKey($planServicePlanId)) {
            $planServicePlanNames.Add($planServicePlanId, $planServicePlanName)
          }
        }

        #get te included service plans
        $srcServicePlanName -split '<br.?>' | ForEach-Object {
          $planProductName = ($_.SubString(0, $_.LastIndexOf('('))).Trim()
          $planServicePlanId = $_.SubString($_.LastIndexOF('(') + 1)
          if ($planServicePlanId.Contains(')')) {
            $planServicePlanId = $planServicePlanId.SubString(0, $planServicePlanId.IndexOf(')'))
          }

          # Add RelevantForTeams
          if ( $planServicePlanNames[$planServicePlanId] ) {
            if ( $planServicePlanNames[$planServicePlanId].Contains('TEAMS') -or $planServicePlanNames[$planServicePlanId].Contains('MCO') ) {
              $Relevant = $true
            }
            else {
              $Relevant = $false
            }
          }
          else {
            $Relevant = $false
          }

          # reworking ProductName into TitleCase
          $VerbosePreference = 'SilentlyContinue'
          $TextInfo = (Get-Culture).TextInfo
          $planProductName = $TextInfo.ToTitleCase($planProductName.ToLower())
          $planProductName = Format-StringRemoveSpecialCharacter -String $planProductName -SpecialCharacterToKeep '()+ -'

          # Building Object
          if ($srcProductPlans.ServicePlanId -notcontains $planServicePlanId) {
            try {
              [void]$srcProductPlans.Add([TFTeamsServicePlan]::new("$planProductName", "$($planServicePlanNames[$planServicePlanId])", "$planServicePlanId", $Relevant))
            }
            catch {
              Write-Debug "[TFTeamsServicePlan] Couldn't add entry for $planProductName"
              if ( $planProductName -ne 'Powerapps For Office 365 K1') {
                $PlansNotAdded += $planProductName
              }
            }
          }
        }
      }
      #endregion

      #region Reworking Parameters
      # Adding ParameterName
      $ParameterName = switch ($srcSkuPartNumber) {
        # Main Licenses
        'M365EDU_A3_FACULTY' { 'Microsoft365A3faculty' }
        'M365EDU_A3_STUDENT' { 'Microsoft365A3students' }
        'M365EDU_A5_FACULTY' { 'Microsoft365A5faculty' }
        'M365EDU_A5_STUDENT' { 'Microsoft365A5students' }
        'SMB_BUSINESS_ESSENTIALS' { 'Microsoft365BusinessBasic' }
        'SMB_BUSINESS_PREMIUM' { 'Microsoft365BusinessStandard' }
        'SPB' { 'Microsoft365BusinessPremium' }
        'SPE_E3' { 'Microsoft365E3' }
        'SPE_E5' { 'Microsoft365E5' }
        'M365_F1' { 'Microsoft365F1' }
        'SPE_F1' { 'Microsoft365F3' }
        'ENTERPRISEPREMIUM_FACULTY' { 'Office365A5faculty' }
        'ENTERPRISEPREMIUM_STUDENT' { 'Office365A5students' }
        'STANDARDPACK' { 'Office365E1' }
        'STANDARDWOFFPACK' { 'Office365E2' }
        'ENTERPRISEPACK' { 'Office365E3' }
        'DEVELOPERPACK' { 'Office365E3Dev' }
        'ENTERPRISEWITHSCAL' { 'Office365E4' }
        'ENTERPRISEPREMIUM' { 'Office365E5' }
        'ENTERPRISEPREMIUM_NOPSTNCONF' { 'Office365E5NoAudioConferencing' }
        'DESKLESSPACK' { 'Office365F1' }
        'SPE_E3_USGOV_DOD' { 'Microsoft365E3USGOVDOD' }
        'SPE_E3_USGOV_GCCHIGH' { 'Microsoft365E3USGOVGCCHIGH' }
        'ENTERPRISEPACK_USGOV_DOD' { 'Office365E3USGOVDOD' }
        'ENTERPRISEPACK_USGOV_GCCHIGH' { 'Office365E3USGOVGCCHIGH' }

        # Standalone Licenses
        'MCOCAP' { 'CommonAreaPhone' }
        'PHONESYSTEM_VIRTUALUSER' { 'PhoneSystemVirtualUser' }
        'MCOSTANDARD' { 'SkypeOnlinePlan2' }

        # Add-on Licenses
        'MCOEV' { 'PhoneSystem' }
        'MCOMEETADV' { 'AudioConferencing' }

        # Microsoft Calling Plans
        'MCOPSTN2' { 'InternationalCallingPlan' }
        'MCOPSTN1' { 'DomesticCallingPlan' }
        'MCOPSTN5' { 'DomesticCallingPlan120' }
        'MCOPSTN_5' { 'DomesticCallingPlan120b' }
        'MCOPSTNC' { 'CommunicationCredits' }

        default { '' }
      }

      # determining LicenseType
      if ( $srcProductPlans.Count -gt 1 ) {
        $LicenseType = 'Package'
        $IncludesTeams = ($srcProductPlans.ServicePlanName -like 'Teams*')
        $IncludesPhoneSystem = ( $srcProductPlans.ServicePlanName -like 'MCOEV*')
      }
      else {
        $LicenseType = switch -Regex ( $srcProductPlans.ServicePlanName ) {
          'MCOPSTN' { 'CallingPlan' }
          'MCOEV' { 'Add-On' }
          'MCOMEETADV' { 'Add-On' }
          default { 'Standalone' }
        }
        $IncludesTeams = ($srcProductPlans.ServicePlanName -like 'Teams*')
        $IncludesPhoneSystem = ( $srcProductPlans.ServicePlanName -like 'MCOEV*')
      }

      # reworking ProductName into TitleCase
      $TextInfo = (Get-Culture).TextInfo
      $ProductName = $TextInfo.ToTitleCase($srcProductName.ToLower())

      #Normalising "SKYPE FOR BUSINESS PSTN" from ProductName for Calling Plans
      $StringToCut = 'Skype For Business Pstn '
      if ( $ProductName -match "$StringToCut" ) {
        $ProductName = $ProductName.Substring($StringToCut.Length, $ProductName.Length - $StringToCut.Length)
      }
      $VerbosePreference = "SilentlyContinue"
      $ProductName = Format-StringRemoveSpecialCharacter -String $ProductName -SpecialCharacterToKeep '()+ '

      # Building Object
      try {
        [void]$Products.Add([TFTeamsLicense]::new( "$ProductName", "$srcSkuPartNumber", "$LicenseType", "$ParameterName", $IncludesTeams, $IncludesPhoneSystem, "$srcSkuId", $srcProductPlans))
      }
      catch {
        Write-Debug "[TFTeamsLicense] Couldn't add entry for '$ProductName'" -Debug
        $ProductsNotAdded += $ProductName
      }
    }

    # Add License Products not on Website
    # Adding Common Area Phone
    [System.Collections.ArrayList]$ServicePlanCAP = @()
    [void]$ServicePlanCAP.Add([TFTeamsServicePlan]::new('Microsoft Teams', 'TEAMS1', '57ff2da0-773e-42df-b2af-ffb7a2317929', $true))
    [void]$ServicePlanCAP.Add([TFTeamsServicePlan]::new('Microsoft 365 Phone System', 'MCOEV', '4828c8ec-dc2e-4779-b502-87ac9ce28ab7', $true))
    [void]$ServicePlanCAP.Add([TFTeamsServicePlan]::new('Skype For Business Online (Plan 2)', 'MCOSTANDARD', '0feaeb32-d00e-4d66-bd5a-43b5b83db82c', $true))
    [void]$Products.Add([TFTeamsLicense]::new( 'Common Area Phone', 'MCOCAP', 'Package', 'CommonAreaPhone', $true, $true, '295a8eb0-f78d-45c7-8b5b-1eed5ed02dff', $ServicePlanCAP))

    # Adding Microsoft 365 Phone System – Virtual User
    [System.Collections.ArrayList]$ServicePlanPSVU = @()
    [void]$ServicePlanPSVU.Add([TFTeamsServicePlan]::new('Phone System - Virtual User', 'MCOEV_VIRTUALUSER', 'f47330e9-c134-43b3-9993-e7f004506889', $true))
    [void]$Products.Add([TFTeamsLicense]::new( 'Microsoft 365 Phone System – Virtual User', 'PHONESYSTEM_VIRTUALUSER', 'Standalone', 'PhoneSystemVirtualUser', $false, $true, '440eaaa8-b3e0-484b-a8be-62870b9ba70a', $ServicePlanPSVU))

    # Adding Communication Credits
    [System.Collections.ArrayList]$ServicePlanCC = @()
    [void]$ServicePlanCC.Add([TFTeamsServicePlan]::new('Communications Credits', 'MCOPSTNC', '505e180f-f7e0-4b65-91d4-00d670bbd18c', $true))
    [void]$Products.Add([TFTeamsLicense]::new( 'Communications Credits', 'MCOPSTNC', 'CallingPlan', 'CommunicationsCredits', $false, $false, '47794cd0-f0e5-45c5-9033-2eb6b5fc84e0', $ServicePlanCC))


    # Output
    if ( $ProductsNotAdded.Count -gt 0 ) {
      Write-Warning -Message "The following Products could not be added: $ProductsNotAdded"
    }

    $ProductsSorted = $Products | Sort-Object ProductName | Sort-Object LicenseType -Desc
    if ($FilterRelevantForTeams) {
      $ProductsSorted = $ProductsSorted | Where-Object { $_.ParameterName -NE '' -or $_.IncludesTeams -or $_.IncludesPhoneSystem }
    }

    return $ProductsSorted

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"

  } #end
} #Get-AzureAdLicense
