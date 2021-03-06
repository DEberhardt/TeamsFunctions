﻿# Module:   TeamsFunctions
# Function: VoiceConfig
# Author:   David Eberhardt
# Updated:  01-DEC-2020
# Status:   Live

#TODO Check for SupportsPaging for OVP and TDP (result size is not managable!)


function Find-TeamsUserVoiceConfig {
  <#
  .SYNOPSIS
    Displays User Accounts matching a specific Voice Configuration Parameter
  .DESCRIPTION
    Returns UserPrincipalNames of Objects matching specific parameters. For PhoneNumbers also displays their basic Voice Configuration
    Search parameters are mutually exclusive, only one Parameter can be specified at the same time.
    Available parameters are:
    - PhoneNumber: Part of the LineURI (ideally without 'tel:','+' or ';ext=...')
    - ConfigurationType: 'CallPlans' or 'DirectRouting'. Will deliver partially configured accounts as well.
    - VoicePolicy: 'BusinessVoice' (CallPlans) or 'HybridVoice' (DirectRouting or any other Hybrid PSTN configuration)
    - OnlineVoiceRoutingPolicy: Any string value (incl. $Null), but not empty ones.
    - TenantDialPlan: Any string value (incl. $Null), but not empty ones.
  .PARAMETER UserPrincipalName
    Optional. UserPrincipalName (UPN) of the User
    Behaves like Get-TeamsUserVoiceConfig, displaying the Users Voice Configuration
  .PARAMETER PhoneNumber
    Optional. Searches all Users matching the given String in their LineURI.
    The expected ResultSize is limited, the full Object is displayed (Get-TeamsUserVoiceConfig)
    Please see NOTES for details
  .PARAMETER ConfigurationType
    Optional. Searches all enabled Users which are at least partially configured for 'CallingPlans', 'DirectRouting' or 'SkypeHybridPSTN'.
    The expected ResultSize is big, therefore only UserPrincipalNames are returned
    Please see NOTES for details
  .PARAMETER VoicePolicy
    Optional. Searches all enabled Users which are reported as 'BusinessVoice' or 'HybridVoice'.
    The expected ResultSize is big, therefore only UserPrincipalNames are returned
    Please see NOTES for details
  .PARAMETER OnlineVoiceRoutingPolicy
    Optional. Searches all enabled Users which have the OnlineVoiceRoutingPolicy specified assigned.
    Please specify full and correct name or '$null' to receive all Users without one
    The expected ResultSize is big, therefore only UserPrincipalNames are returned
    Please see NOTES for details
  .PARAMETER TenantDialPlan
    Optional. Searches all enabled Users which have the TenantDialPlan specified assigned.
    Please specify full and correct name or '$null' to receive all Users without one
    The expected ResultSize is big, therefore only UserPrincipalNames are returned
    Please see NOTES for details
  .PARAMETER ValidateLicense
    Optional. Can be combined only with -ConfigurationType
    In addition to validation of Parameters, also validates License assignment for the found user.
    License Check is performed AFTER parameters are verified.
  .EXAMPLE
    Find-TeamsUserVoiceConfig -UserPrincipalName John@domain.com
    Shows Voice Configuration for John, returning the full Object
  .EXAMPLE
    Find-TeamsUserVoiceConfig -PhoneNumber "15551234567"
    Shows all Users which have this String in their LineURI (TelephoneNumber or OnPremLineURI)
    The expected ResultSize is limited, the full Object is returned (Get-TeamsUserVoiceConfig)
    Please see NOTES for details
  .EXAMPLE
    Find-TeamsUserVoiceConfig -ConfigurationType CallingPlans
    Shows all Users which are configured for CallingPlans (Full)
    The expected ResultSize is big, therefore only Names (UPNs) of Users are returned
    Pipe to Get-TeamsUserVoiceConfiguration for full output.
    Please see NOTES for details
  .EXAMPLE
    Find-TeamsUserVoiceConfig -VoicePolicy BusinessVoice
    Shows all Users which are configured for PhoneSystem with CallingPlans
    The expected ResultSize is big, therefore only Names (UPNs) of Users are displayed
    Pipe to Get-TeamsUserVoiceConfiguration for full output.
    Please see NOTES and LINK for details
  .EXAMPLE
    Find-TeamsUserVoiceConfig -OnlineVoiceRoutingPolicy O_VP_EMEA
    Shows all Users which have the OnlineVoiceRoutingPolicy "O_VP_EMEA" assigned
    The expected ResultSize is big, therefore only Names (UPNs) of Users are displayed
    Pipe to Get-TeamsUserVoiceConfiguration for full output.
    Please see NOTES for details
  .EXAMPLE
    Find-TeamsUserVoiceConfig -TenantDialPlan DP-US
    Shows all Users which have the TenantDialPlan "DP-US" assigned.
    Please see NOTES for details
  .INPUTS
    System.String
  .OUTPUTS
    System.String - UserPrincipalName - With any Parameter except Identity or PhoneNumber
    System.Object - With Parameter Identity or PhoneNumber
  .NOTES
    With the exception of Identity and PhoneNumber, all searches are filtering on Get-CsOnlineUser
    This usually should not take longer than a minute to complete.
    Identity is querying the provided UPN and only wraps Get-TeamsUserVoiceConfig
    PhoneNumber has to do a full search with 'Where-Object' which will take time to complete
    Depending on the number of Users in the Tenant, this may take a few minutes!

    All Parameters except UserPrincipalName or PhoneNumber will only return UserPrincipalNames (UPNs)
    - PhoneNumber: Searches against the LineURI parameter. For best compatibility, provide in E.164 format (with or without the +)
    This script can find duplicate assignments if the Number was assigned with and without an extension.
    - ConfigurationType: This is determined with Test-TeamsUserVoiceConfig -Partial and will return all Accounts found
    - VoicePolicy: BusinessVoice are PhoneSystem Users exclusively configured for Microsoft Calling Plans.
      HybridVoice are PhoneSystem Users who are configured for TDR, Hybrid SkypeOnPrem PSTN or Hybrid CloudConnector PSTN breakouts
    - OnlineVoiceRoutingPolicy: Finds all users which have this particular Policy assigned
    - TenantDialPlan: Finds all users which have this particular DialPlan assigned.
    Please see Related Link for more information
  .COMPONENT
    VoiceConfiguration
  .FUNCTIONALITY
    Finding Users with a specific values in their Voice Configuration
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Find-TeamsUserVoiceConfig.md
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_VoiceConfiguration.md
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_UserManagement.md
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    https://docs.microsoft.com/en-us/microsoftteams/direct-routing-migrating
  #>

  [CmdletBinding(DefaultParameterSetName = 'Tel')]
  [Alias('Find-TeamsUVC')]
  [OutputType([PSCustomObject])]
  param(
    [Parameter(ParameterSetName = 'ID')]
    [Alias('ObjectId', 'Identity')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Tel', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'String to be found in any of the PhoneNumber fields')]
    [ValidateScript( {
        If ($_ -match '^(tel:\+|\+)?([0-9]?[-\s]?(\(?[0-9]{3}\)?)[-\s]?([0-9]{3}[-\s]?[0-9]{4})|([0-9][-\s]?){4,20})((x|;ext=)([0-9]{3,8}))?$') { $True } else {
          throw [System.Management.Automation.ValidationMetadataException] 'Not a valid phone number format. Expected min 4 digits, but multiple formats accepted. Extensions will be stripped'
          $false
        }
      })]
    [Alias('Number', 'TelephoneNumber', 'Tel', 'LineURI', 'OnPremLineURI')]
    [string[]]$PhoneNumber,

    [Parameter(ParameterSetName = 'Ext', HelpMessage = 'String to be found in any of the PhoneNumber fields as an Extension')]
    [Alias('Ext')]
    [string[]]$Extension,

    [Parameter(ParameterSetName = 'CT', HelpMessage = 'Filters based on Configuration Type')]
    [ValidateSet('CallingPlans', 'SkypeHybridPSTN', 'DirectRouting')]
    [String]$ConfigurationType,

    [Parameter(ParameterSetName = 'VP', HelpMessage = 'Filters based on VoicePolicy')]
    [ValidateSet('BusinessVoice', 'HybridVoice')]
    [String]$VoicePolicy,

    [Parameter(ParameterSetName = 'OVP', HelpMessage = 'Filters based on OnlineVoiceRoutingPolicy')]
    [AllowNull()]
    [Alias('OVP')]
    [String]$OnlineVoiceRoutingPolicy,

    [Parameter(ParameterSetName = 'TDP', HelpMessage = 'Filters based on TenantDialPlan')]
    [AllowNull()]
    [Alias('TDP')]
    [String]$TenantDialPlan,

    [Parameter(ParameterSetName = 'CT', HelpMessage = 'Additionally also validates License (CallingPlan or PhoneSystem)')]
    [switch]$ValidateLicense

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
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

    if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
      Write-Warning -Message "The switch 'ValidateLicense' verifies whether the correct license is assigned before considering the User. This increases run-time tremendously!"
    }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    switch ($PsCmdlet.ParameterSetName) {
      'ID' {
        Write-Information "Finding Users with SipAddress '$UserPrincipalName'"
        #Filter must be written as-is (Get-CsOnlineUser is an Online command, handover of parameters is sketchy)
        $Filter = 'SipAddress -like "*{0}*"' -f $UserPrincipalName
        $Users = Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object UserPrincipalName
        if (-not $Users) {
          $MailNickName = $UserPrincipalName.split('@') | Select-Object -First 1
          Write-Information "Finding Users with MailNickName '$MailNickName'"
          $Filter = 'MailNickName -like "*{0}*"' -f $MailNickName
          $Users = Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object UserPrincipalName
        }

        if ($Users) {
          if ($Users.Count -gt 3) {
            Write-Verbose -Message 'Multiple results found - Displaying limited output only' -Verbose
            $Users | Select-Object UserPrincipalName, TelephoneNumber, LineUri, OnPremLineURI
          }
          else {
            Write-Verbose -Message 'Limited results found - Displaying User Voice Configuration for each' -Verbose
            Get-TeamsUserVoiceConfig -UserPrincipalName $($Users.UserPrincipalName)
          }
        }
        else {
          Write-Verbose -Message "User: '$UserPrincipalName' - No records found (SipAddress)" -Verbose
        }
        break
      } #ID

      'Tel' {
        foreach ($PhoneNr in $PhoneNumber) {
          Write-Verbose -Message "Normalising Input for Phone Number '$PhoneNr'"
          if ($PhoneNr -match '@') {
            Find-TeamsUserVoiceConfig -UserPrincipalName "$PhoneNr"
            continue
          }
          elseif ($PhoneNr -match '([0-9]{3,25});ext=([0-9]{3,8})') {
            $Number = $matches[1] # Phone Number
            # $Number = $matches[2] # Extension
          }
          else {
            $Number = Format-StringForUse "$($PhoneNr.split(';')[0].split('x')[0])" -SpecialChars 'telx:+() -'
            # Alternative (working but not stripping/capturing extensions with x)
            #$Number = Format-StringRemoveSpecialCharacter "$PhoneNr" | Format-StringForUse -SpecialChars 'tel'
          }
          Write-Information "Finding all Users enabled for Teams with Phone Number string '$Number': Searching..."
          #Filter must be written as-is (Get-CsOnlineUser is an Online command, handover of parameters is sketchy)
          $Filter = 'LineURI -like "*{0}*"' -f $Number
          $Users = Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
          if ($Users) {
            if ($Users.Count -gt 1) {
              Write-Warning -Message "Number: '$Number' - Found multiple Users matching the criteria! If the search string represents the FULL number, it is assigned incorrectly. Inbound calls to this number will not work as Teams will not find a unique match"
              Write-Verbose -Message "Investigate OnPremLineURI string. Has one of them set an Extension (';ext=') set, the other one not?" -Verbose
            }
            if ($Users.Count -gt 3) {
              Write-Verbose -Message 'Multiple results found - Displaying limited output only' -Verbose
              $Users | Select-Object UserPrincipalName, LineUri
            }
            else {
              Write-Verbose -Message 'Limited results found - Displaying User Voice Configuration for each' -Verbose
              Get-TeamsUserVoiceConfig -UserPrincipalName $($Users.UserPrincipalName)
            }
          }
          else {
            Write-Verbose -Message "Number: '$Number' - No assignments found (LineURI)" -Verbose
          }
        }
        break
      } #Tel

      'Ext' {
        foreach ($Ext in $Extension) {
          Write-Verbose -Message "Normalising Input for Extension '$Ext'"
          if ($ext -match '([0-9]{3,15})?;?ext=([0-9]{3,8})') {
            # $Number = $matches[1] # Phone Number
            # $Number = $matches[2] # Extension
            $ExtN = 'ext=' + $matches[2]
          }
          else {
            $ExtN = 'ext=' + $ext
          }
          Write-Information "Finding all Users enabled for Teams with Extension '$ExtN': Searching..."
          #Filter must be written as-is (Get-CsOnlineUser is an Online command, handover of parameters is sketchy)
          $Filter = 'LineURI -like "*{0}*"' -f "$ExtN"
          $Users = Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
          if ($Users) {
            if ($Users.Count -gt 1) {
              Write-Warning -Message "Extension: '$ExtN' - Found multiple Users matching the criteria! If the search string represents the FULL extension, it is assigned incorrectly. Inbound calls to this extension may fail depending on normalisation as Teams will not find a unique match"
              Write-Verbose -Message 'Investigate OnPremLineURI string. Verify unique Extension is applied.' -Verbose
            }
            if ($Users.Count -gt 3) {
              Write-Verbose -Message 'Multiple results found - Displaying limited output only' -Verbose
              $Users | Select-Object UserPrincipalName, LineUri
            }
            else {
              Write-Verbose -Message 'Limited results found - Displaying User Voice Configuration for each' -Verbose
              Get-TeamsUserVoiceConfig -UserPrincipalName $($Users.UserPrincipalName)
            }
          }
          else {
            Write-Verbose -Message "Extension: '$ExtN' - No assignments found (LineURI)" -Verbose
          }
        }
        break
      } #Ext

      'CT' {
        Write-Verbose -Message 'Searching for all Users enabled for Teams: Searching... This will take quite some time!'
        $Filter = 'Enabled -eq $TRUE'
        $CsUsers = Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Verbose -Message "Sifting through Information for $($CsUsers.Count) Users: Parsing..."
        Write-Information "Finding all Users enabled for Teams with ConfigurationType '$ConfigurationType' Searching... This will take quite some time!"
        switch ($ConfigurationType) {
          'DirectRouting' {
            if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
              Write-Verbose -Message 'Switch ValidateLicense: Only users with PhoneSystem license (enabled ServicePlan) are displayed!' -Verbose
            }
            #CHECK Revisit this based on Test-TeamsUserVoiceConfig
            foreach ($U in $CsUsers) {
              if ($U.VoicePolicy -eq 'HybridVoice' -and $null -eq $U.VoiceRoutingPolicy -and ($null -ne $U.OnPremLineURI -or $null -ne $U.OnlineVoiceRoutingPolicy)) {
                if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
                  if (Test-TeamsUserLicense $U -ServicePlan MCOEV) {
                    $U.UserPrincipalName
                  }
                }
                else {
                  $U.UserPrincipalName
                }
              }
            }
            break
          }
          'SkypeHybridPSTN' {
            if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
              Write-Verbose -Message 'Switch ValidateLicense: Only users with PhoneSystem license (enabled ServicePlan) are displayed!' -Verbose
            }
            foreach ($U in $CsUsers) {
              if ($U.VoicePolicy -eq 'HybridVoice' -and $null -eq $U.OnlineVoiceRoutingPolicy -and ($null -ne $U.OnPremLineURI -or $null -ne $U.VoiceRoutingPolicy)) {
                if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
                  if (Test-TeamsUserLicense $U -ServicePlan MCOEV) {
                    $U.UserPrincipalName
                  }
                }
                else {
                  $U.UserPrincipalName
                }
              }
              break
            }
          }
          'CallingPlans' {
            if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
              Write-Verbose -Message 'Switch ValidateLicense: Only users with CallPlan license are displayed!' -Verbose
            }
            foreach ($U in $CsUsers) {
              if ($U.VoicePolicy -eq 'BusinessVoice' -or $null -ne $U.TelephoneNumber) {
                if ($PSBoundParameters.ContainsKey('ValidateLicense')) {
                  if (Test-TeamsUserHasCallPlan $U) {
                    $U.UserPrincipalName
                  }
                }
                else {
                  $U.UserPrincipalName
                }
              }
            }
            break
          }
        }
        break
      } #CT

      'VP' {
        Write-Information "Finding all Users enabled for Teams with VoicePolicy '$VoicePolicy': Searching... This will take a bit of time!"
        $Filter = 'Enabled -eq $TRUE -and  VoicePolicy -EQ "{0}"' -f $VoicePolicy
        Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue | Select-Object UserPrincipalName
        break
      } #VP

      'OVP' {
        Write-Verbose -Message "Finding OnlineVoiceRoutingPolicy '$OnlineVoiceRoutingPolicy'..."
        $OVP = Get-CsOnlineVoiceRoutingPolicy $OnlineVoiceRoutingPolicy -WarningAction SilentlyContinue
        if ($null -ne $OVP) {
          Write-Information "Finding all Users enabled for Teams with OnlineVoiceRoutingPolicy '$OnlineVoiceRoutingPolicy': Searching... This will take a bit of time!"
          $Filter = 'Enabled -eq $TRUE -and  OnlineVoiceRoutingPolicy -EQ "{0}"' -f $OnlineVoiceRoutingPolicy
          Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue | Select-Object UserPrincipalName
        }
        else {
          Write-Error -Message "OnlineVoiceRoutingPolicy '$OnlineVoiceRoutingPolicy' not found" -Category ObjectNotFound -ErrorAction Stop
        }
        break
      } #OVP

      'TDP' {
        Write-Verbose -Message "Finding TenantDialPlan '$TenantDialPlan'..."
        $TDP = Get-CsTenantDialPlan $TenantDialPlan -WarningAction SilentlyContinue
        if ($null -ne $TDP) {
          Write-Information "Finding all Users enabled for Teams with TenantDialPlan '$TenantDialPlan': Searching... This will take a bit of time!"
          $Filter = 'Enabled -eq $TRUE -and  TenantDialPlan -EQ "{0}"' -f $TenantDialPlan
          Get-CsOnlineUser -Filter $Filter -WarningAction SilentlyContinue | Select-Object UserPrincipalName
        }
        else {
          Write-Error -Message "TenantDialPlan '$TenantDialPlan' not found" -Category ObjectNotFound -ErrorAction Stop
        }
        break
      } #TDP

      default {
        # No Parameter is specified
        Write-Warning -Message 'No Parameters specified. Please specify search criteria (Parameter and value)!' -Verbose
        break
      } #default

    } #Switch

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} # Find-TeamsUserVoiceConfig
