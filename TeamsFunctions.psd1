#
# Module manifest for module 'TeamsFunctions'
#
# Generated by: David Eberhardt
#
# Generated on: 16/05/2020
#

@{

  # Script module or binary module file associated with this manifest.
  RootModule            = 'TeamsFunctions.psm1'

  # Version number of this module.
  ModuleVersion         = '20.12.20'

  # Supported PSEditions
  # CompatiblePSEditions = @()

  # ID used to uniquely identify this module
  GUID                  = 'c0165b45-500b-4d59-be47-64f567e4b4c9'

  # Author of this module
  Author                = 'David Eberhardt'

  # Company or vendor of this module
  CompanyName           = 'None / Personal'

  # Copyright statement for this module
  Copyright             = '(c) 2020 David Eberhardt. All Rights Reserved'

  # Description of the functionality provided by this module
  Description           = 'Teams Functions for Administration of Users, Resource Accounts, Call Queues and Auto Attendants, incl. Licensing, User Voice Configuration with Calling Plans and Direct Routing,
For more information, please visit the https://davideberhardt.wordpress.com/ or https://github.com/DEberhardt/TeamsFunctions'

  # Minimum version of the Windows PowerShell engine required by this module
  PowerShellVersion     = '5.1'

  # Name of the Windows PowerShell host required by this module
  # PowerShellHostName = ''

  # Minimum version of the Windows PowerShell host required by this module
  # PowerShellHostVersion = ''

  # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
  # DotNetFrameworkVersion = ''

  # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
  # CLRVersion = ''

  # Processor architecture (None, X86, Amd64) required by this module
  ProcessorArchitecture = 'Amd64'

  # Modules that must be imported into the global environment prior to importing this module
  RequiredModules       = @('AzureAdPreview')
  #RequiredModules = @('AzureAd')

  # Assemblies that must be loaded prior to importing this module
  # RequiredAssemblies = @()

  # Script files (.ps1) that are run in the caller's environment prior to importing this module.
  # ScriptsToProcess = @()

  # Type files (.ps1xml) to be loaded when importing this module
  # TypesToProcess = @()

  # Format files (.ps1xml) to be loaded when importing this module
  # FormatsToProcess = @()

  # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
  # NestedModules = @()

  # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
  FunctionsToExport     = @(
    # Auto Attendant
    'New-TeamsAutoAttendant',
    'Get-TeamsAutoAttendant',
    'Remove-TeamsAutoAttendant',
    'New-TeamsAutoAttendantCallFlow',
    'New-TeamsAutoAttendantMenu',
    'New-TeamsAutoAttendantMenuOption',
    'New-TeamsAutoAttendantPrompt',
    'New-TeamsAutoAttendantSchedule',
    'New-TeamsAutoAttendantDialScope',

    # Call Queue
    'New-TeamsCallQueue',
    'Get-TeamsCallQueue',
    'Set-TeamsCallQueue',
    'Remove-TeamsCallQueue',

    # Licensing
    'Add-TeamsUserLicense',
    'Set-TeamsUserLicense',
    'Get-TeamsUserLicense',
    'Get-TeamsTenantLicense',
    'Get-TeamsLicense',
    'Get-TeamsLicenseServicePlan',
    'Get-AzureAdLicense',
    'Get-AzureAdLicenseServicePlan',

    # Resource Account
    'New-TeamsResourceAccount',
    'Get-TeamsResourceAccount',
    'Find-TeamsResourceAccount',
    'Set-TeamsResourceAccount',
    'Remove-TeamsResourceAccount',
    'New-TeamsResourceAccountAssociation',
    'Get-TeamsResourceAccountAssociation',
    'Remove-TeamsResourceAccountAssociation',

    # Session
    'Assert-AzureADConnection',
    'Assert-SkypeOnlineConnection',
    'Assert-MicrosoftTeamsConnection',
    'Connect-SkypeOnline',
    'Disconnect-SkypeOnline',
    'Connect-Me',
    'Disconnect-Me',
    'Test-AzureADConnection',
    'Test-SkypeOnlineConnection',
    'Test-MicrosoftTeamsConnection',
    'Test-ExchangeOnlineConnection',

    # Support
    'Backup-TeamsEV',
    'Restore-TeamsEV',
    'Backup-TeamsTenant',

    'Format-StringForUse',
    'Format-StringRemoveSpecialCharacter',
    'Test-AzureAdGroup',
    'Test-AzureAdUser',
    'Test-TeamsResourceAccount'
    'Test-Module',
    'Test-TeamsUser',
    'Get-TeamsObjectType',
    'Write-ErrorRecord',

    'Get-TeamsMGW',
    'Get-TeamsOVR',
    'Get-TeamsOPU',
    'Get-TeamsOVP',
    'Get-TeamsTDP',
    'Get-TeamsVNR',
    'Get-TeamsTenant',

    'Get-RegionFromCountryCode',
    'Get-SkuPartNumberfromSkuId',
    'Get-SkuIdFromSkuPartNumber',
    'New-AzureAdLicenseObject',
    'Test-TeamsUserLicense',
    'Test-TeamsUserHasCallPlan',

    'Get-SkypeOnlineConferenceDialInNumbers',

    'Import-TeamsAudioFile',
    'Remove-TenantDialPlanNormalizationRule',
    'Set-TeamsUserPolicy',
    'Test-TeamsExternalDNS',
    'Test-TeamsTenantPolicy',


    #UserManagement
    'Get-TeamsCallableEntity',
    'Find-TeamsCallableEntity',
    'New-TeamsCallableEntity',
    'Assert-TeamsCallableEntity',
    'Find-AzureAdUser',
    'Find-AzureAdGroup',

    #UserAdmin
    'Get-AzureAdAssignedAdminRoles',
    'Enable-AzureAdAdminRole',

    # Voice Config
    'Get-TeamsTenantVoiceConfig',
    'Get-TeamsUserVoiceConfig',
    'Find-TeamsUserVoiceConfig',
    'Set-TeamsUserVoiceConfig',
    'Remove-TeamsUserVoiceConfig',
    'Test-TeamsUserVoiceConfig',
    'Enable-TeamsUserForEnterpriseVoice'#,

    #'Get-TeamsVoiceRoutingConfig',
    #'New-TeamsCommonAreaPhone',
    #'Remove-TeamsCommonAreaPhone'
  )

  # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
  CmdletsToExport       = @()

  # Variables to export from this module
  #VariablesToExport     = @('TeamsLicenses', 'TeamsServicePlans')

  # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
  AliasesToExport       = @(
    'con',
    'dis',
    'pol',

    'Set-TeamsUVC',
    'Get-TeamsUVC',
    'Find-TeamsUVC',
    'Remove-TeamsUVC',
    'Test-TeamsUVC',
    'Enable-Ev',
    #'New-TeamsCAP',
    #'Remove-TeamsCAP',

    'New-TeamsRA',
    'Set-TeamsRA',
    'Get-TeamsRA',
    'Find-TeamsRA',
    'Remove-TeamsRA',
    'New-TeamsRAassoc',
    'Get-TeamsRAassoc',
    'Remove-TeamsRAassoc',
    'Remove-CsOnlineApplicationInstance',

    'New-TeamsCQ',
    'Set-TeamsCQ',
    'Get-TeamsCQ',
    'Remove-TeamsCQ',

    'New-TeamsAA',
    'Set-TeamsAA',
    'Get-TeamsAA',
    'Remove-TeamsAA',
    'Set-TeamsAutoAttendant',

    'New-TeamsAAMenu',
    'New-TeamsAAOption',
    'New-TeamsAAFlow',
    'New-TeamsAAScope',
    'New-TeamsAAPrompt',
    'New-TeamsAASchedule',
    'New-TeamsAAEntity',
    'New-TeamsAutoAttendantCallableEntity'
    )

  # DSC resources to export from this module
  # DscResourcesToExport = @()

  # List of all modules packaged with this module
  # ModuleList = @()

  # List of all files packaged with this module
  # FileList = @()

  # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
  PrivateData           = @{

    PSData = @{

      # Tags applied to this module. These help with module discovery in online galleries.
      Tags       = @('PowerShell', 'Teams', 'DirectRouting', 'SkypeOnline', 'Licensing', 'ResourceAccount', 'CallQueue', 'AutoAttendant', 'VoiceConfig')

      # Prerelease Version
      Prerelease = '-prerelease'

      # A URL to the license for this module.
      # LicenseUri = ''

      # A URL to the main website for this project.
      # ProjectUri = ''

      # A URL to an icon representing this module.
      # IconUri = ''

      # ReleaseNotes of this module
      # ReleaseNotes = ''

    } # End of PSData hashtable

  } # End of PrivateData hashtable

  # HelpInfo URI of this module
  # HelpInfoURI = ''

  # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
  # DefaultCommandPrefix = ''

}
