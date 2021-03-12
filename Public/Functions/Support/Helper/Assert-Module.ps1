# Module:   TeamsFunctions
# Function: Support
# Author:		David Eberhardt
# Updated:  01-JUL-2020
# Status:   Live

function Assert-Module {
  <#
	.SYNOPSIS
		Tests whether a Module is loaded
  .DESCRIPTION
    Tests whether a specific Module is loaded
  .PARAMETER Module
    Names of one or more Modules to assert
  .PARAMETER UpToDate
    Verifies Version installed is equal to the latest found online
  .PARAMETER PreRelease
    Verifies Version installed is equal to the latest prerelease version found online
  .EXAMPLE
		Assert-Module -Module ModuleName
		Will Return $TRUE if the Module 'ModuleName' is installed and loaded
  .EXAMPLE
		Assert-Module -Module ModuleName -UpToDate
		Will Return $TRUE if the Module 'ModuleName' is installed in the latest release version and loaded
  .EXAMPLE
		Assert-Module -Module ModuleName -UpToDate -PreRelease
		Will Return $TRUE if the Module 'ModuleName' is installed in the latest pre-release version and loaded
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  #>

  [CmdletBinding()]
  [OutputType([Boolean])]
  Param
  (
    [Parameter(Position = 0, HelpMessage = 'Module to test')]
    [string[]]$Module,

    [Parameter(HelpMessage = 'Verifies the latest version is installed')]
    [switch]$UpToDate,

    [Parameter(HelpMessage = 'Verifies the latest prerelease version is installed')]
    [switch]$PreRelease

  )

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    foreach ($M in $Module) {
      Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Checking Installation"
      $Installed = Get-Module -Name $M -ListAvailable -Verbose:$false -WarningAction SilentlyContinue
      if ( -not $Installed) {
        Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Checking Installation - Module not installed"
        return $false
      }
      else {
        # Determining Current Version
        if ($Installed.count -gt 1) { $Current = $Installed[0] } else { $Current = $Installed }
        Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Current Version installed: $($Current.Version)"
        $CurrentVersion = [Version] ($Current.Version.ToString() -replace '^(\d+(\.\d+){0,3})(\.\d+?)*$' , '$1')
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name): CurrentVersion:", ($CurrentVersion | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }

        if ($UpToDate) {
          # Checking Current Version is UpToDate
          Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Checking Version"
          $FindModuleParams = $null
          $FindModuleParams += @{'Name' = $M }
          $FindModuleParams += @{'Verbose' = $false }
          $FindModuleParams += @{'Debug' = $false }
          $FindModuleParams += @{'ErrorAction' = 'SilentlyContinue' }
          if ($PreRelease) { $FindModuleParams += @{'AllowPrerelease' = $true } }
          $Latest = Find-Module @FindModuleParams
          $LatestVersion = if ($Latest.Version -match '-') { $Latest.Version.Split('-')[0] } else { $Latest.Version }
          $LatestVersion = [Version] ($LatestVersion.ToString() -replace '^(\d+(\.\d+){0,3})(\.\d+?)*$' , '$1')
          if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
            "Function: $($MyInvocation.MyCommand.Name): LatestVersion:", ($LatestVersion | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
          }

          if ($CurrentVersion -lt $LatestVersion) {
            # $CurrentVersion is less than $LatestVersion
            Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Update available! Latest Version: $($Latest.Version)" -Verbose
            return $false
          }
        }

        # Checking Imported Version is CurrentVersion
        Write-Verbose -Message "$($MyInvocation.MyCommand) - Verifying Module '$M' - Checking Import"
        $CurrentlyLoaded = Get-Module -Name $M -Verbose:$false -WarningAction SilentlyContinue
        if ($null -ne $CurrentlyLoaded) {
          $CurrentlyLoadedVersion = [Version] ($CurrentlyLoaded.Version.ToString() -replace '^(\d+(\.\d+){0,3})(\.\d+?)*$' , '$1')
          if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
            "Function: $($MyInvocation.MyCommand.Name): CurrentlyLoadedVersion:", ($CurrentlyLoadedVersion | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
          }
          if ($CurrentlyLoadedVersion -ne $CurrentVersion ) {
            Write-Verbose -Message "Removing Module '$M' - Version $CurrentlyLoadedVersion"
            Remove-Module -Name $M -Force -Verbose:$false -ErrorAction SilentlyContinue
          }
        }
        $Loaded = Get-Module -Name $M -Verbose:$false -WarningAction SilentlyContinue
        if ($null -ne $Loaded) {
          return $true
        }
        else {
          Write-Verbose -Message "Importing Module '$M' - Version $CurrentVersion"
          $SaveVerbosePreference = $global:VerbosePreference;
          $global:VerbosePreference = 'SilentlyContinue';
          Import-Module -Name $M -Global -Force -Verbose:$false -ErrorAction SilentlyContinue
          $global:VerbosePreference = $SaveVerbosePreference
          if (Get-Module -Name $M -WarningAction SilentlyContinue) {
            return $true
          }
          else {
            return $false
          }
        }
      }
    }
  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"

  } #end
} # Assert-Module
