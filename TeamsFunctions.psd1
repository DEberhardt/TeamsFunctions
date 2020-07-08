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
  ModuleVersion         = '20.7.8'

  # Supported PSEditions
  # CompatiblePSEditions = @()

  # ID used to uniquely identify this module
  GUID                  = 'c0165b45-500b-4d59-be47-64f567e4b4c9'

  # Author of this module
  Author                = 'David Eberhardt', 'Jeff Brown', 'Ken Lasko'

  # Company or vendor of this module
  CompanyName           = 'None / Personal'

  # Copyright statement for this module
  Copyright             = '(c) 2020 David Eberhardt, Jeff Brown, Ken Lasko. All Rights Reserved'

  # Description of the functionality provided by this module
  Description           = 'Teams Functions for Administration of SkypeOnline (Teams Backend), Voice and Direct Routing, Resource Accounts, Call Queues, etc.

Expanding on functionality for AzureAD, adding queries for Objects by friendly Names (UPNs) in AzureAD, Licensing and Test Cmdlets

Expanding on functionality for SkypeOnline, improving Session management and connection, Tenant Backup (curtesy of Ken Lasko), etc.
Improved Resource Account Management: Search and select by friendly names and UPNs,
Improved Call Queue Management: Search and select by friendly names and UPNs.

Updated regularly to add functionality as required.
For more information, please visit the following: https://davideberhardt.wordpress.com/

Connect via GitHub: https://github.com/DEberhardt/TeamsFunctions'

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
  RequiredModules       = @('AzureAd', 'MicrosoftTeams')
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
  FunctionsToExport     = @('Connect-SkypeOnline', 'Disconnect-SkypeOnline', 'Connect-SkypeTeamsAndAAD', 'Disconnect-SkypeTeamsAndAAD', `
      'Get-AzureAdAssignedAdminRoles', 'Get-AzureADUserFromUPN', `
      'Add-TeamsUserLicense', 'New-AzureAdLicenseObject', 'Get-TeamsUserLicense', 'Get-TeamsTenantLicenses', `
      'Test-TeamsUserLicense', 'Set-TeamsUserPolicy', 'Test-TeamsTenantPolicy', `
      'Test-AzureADModule', 'Test-AzureADConnection', 'Test-AzureADUser', 'Test-AzureADGroup', `
      'Test-SkypeOnlineModule', 'Test-SkypeOnlineConnection', `
      'Test-MicrosoftTeamsModule', 'Test-MicrosoftTeamsConnection', 'Test-TeamsUser', `
      'New-TeamsResourceAccount', 'Get-TeamsResourceAccount', 'Find-TeamsResourceAccount', 'Set-TeamsResourceAccount', 'Remove-TeamsResourceAccount', `
      'New-TeamsResourceAccountAssociation', 'Get-TeamsResourceAccountAssociation', 'Remove-TeamsResourceAccountAssociation', `
      'New-TeamsCallQueue', 'Get-TeamsCallQueue', 'Set-TeamsCallQueue', 'Remove-TeamsCallQueue', `
      'Import-TeamsAudioFile', 'Backup-TeamsEV', 'Restore-TeamsEV', 'Backup-TeamsTenant', `
      'Remove-TenantDialPlanNormalizationRule', 'Test-TeamsExternalDNS', 'Get-SkypeOnlineConferenceDialInNumbers', `
      'Get-SkuPartNumberfromSkuID', 'Get-SkuIDfromSkuPartNumber', 'Format-StringRemoveSpecialCharacter', 'Format-StringForUse', 'Write-ErrorRecord')

  # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
  CmdletsToExport       = @('Connect-SkypeOnline', 'Disconnect-SkypeOnline', 'Get-AzureAdAssignedAdminRoles', 'Get-AzureADUserFromUPN', `
      'Add-TeamsUserLicense', 'New-AzureAdLicenseObject', 'Get-TeamsUserLicense', 'Get-TeamsTenantLicenses', `
      'Test-TeamsUserLicense', 'Set-TeamsUserPolicy', 'Test-TeamsTenantPolicy', `
      'Test-AzureADModule', 'Test-AzureADConnection', 'Test-AzureADUser', 'Test-AzureADGroup', `
      'Test-SkypeOnlineModule', 'Test-SkypeOnlineConnection', `
      'Test-MicrosoftTeamsModule', 'Test-MicrosoftTeamsConnection', 'Test-TeamsUser', `
      'New-TeamsResourceAccount', 'Get-TeamsResourceAccount', 'Find-TeamsResourceAccount', 'Set-TeamsResourceAccount', 'Remove-TeamsResourceAccount', `
      'New-TeamsResourceAccountAssociation', 'Get-TeamsResourceAccountAssociation', 'Remove-TeamsResourceAccountAssociation', `
      'New-TeamsCallQueue', 'Get-TeamsCallQueue', 'Set-TeamsCallQueue', 'Remove-TeamsCallQueue', `
      'Import-TeamsAudioFile', 'Backup-TeamsEV', 'Restore-TeamsEV', 'Backup-TeamsTenant', `
      'Remove-TenantDialPlanNormalizationRule', 'Test-TeamsExternalDNS', 'Get-SkypeOnlineConferenceDialInNumbers', `
      'Get-SkuPartNumberfromSkuID', 'Get-SkuIDfromSkuPartNumber', 'Format-StringRemoveSpecialCharacter', 'Format-StringForUse', 'Write-ErrorRecord')

  # Variables to export from this module
  VariablesToExport     = '*'

  # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
  AliasesToExport       = '*'

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
      Tags       = @('PowerShell', 'Teams', 'DirectRouting', 'SkypeOnline', 'Licensing', 'ResourceAccount', 'CallQueue')

      # Prerelease Version
      Prerelease = '-alpha2'

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
