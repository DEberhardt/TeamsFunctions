# Module:   TeamsFunctions
# Function: Session
# Author:		David Eberhardt
# Updated:  01-JAN-2021
# Status:   Live




function Connect-Me {
  <#
	.SYNOPSIS
		Connect to AzureAd, MicrosoftTeams and optionally also to Exchange
	.DESCRIPTION
		One function to connect them all.
    This CmdLet solves the requirement for individual authentication prompts for AzureAD and MicrosoftTeams
    (and optionally also to ExchangeOnline) when multiple connections are required.
	.PARAMETER AccountId
		Required. UserPrincipalName or LoginName of the Office365 Administrator
	.PARAMETER ExchangeOnline
		Optional. Connects to Exchange Online Management. Requires Exchange Admin Role
	.PARAMETER NoFeedback
		Optional. Suppresses output session information about established sessions. Used for calls by other functions
	.EXAMPLE
		Connect-Me [-AccountId] admin@domain.com
    Creates a session to AzureAD prompting for a Password for 'admin@domain.com'
    If AzureAdPreview is loaded, tries to enable eligible Admin roles in Privileged Identity Management
    Creates a session to MicrosoftTeams with the AzureAd Session details
    If unsuccessful, prompting for selection of the authenticated User only (no additional authentication needed)
	.EXAMPLE
		Connect-Me -AccountId admin@domain.com -NoFeedBack
    If AzureAdPreview is loaded, tries to enable eligible Admin roles in Privileged Identity Management
    Creates a session to MicrosoftTeams with the AzureAd Session details
    If unsuccessful, prompting for selection of the authenticated User only (no additional authentication needed)
    Does not display Session Information Object at the end - This is useful if called by other functions.
	.EXAMPLE
		Connect-Me -AccountId admin@domain.com -ExchangeOnline
    If AzureAdPreview is loaded, tries to enable eligible Admin roles in Privileged Identity Management
    Creates a session to MicrosoftTeams with the AzureAd Session details
    If unsuccessful, prompting for selection of the authenticated User only (no additional authentication needed)
    Also connects to ExchangeOnline
  .FUNCTIONALITY
    Connects to one or multiple Office 365 Services with as few Authentication prompts as possible
  .NOTES
    This CmdLet can be used to establish a session to: AzureAD, MicrosoftTeams and ExchangeOnline
    Each Service has different requirements for connection, query (Get-CmdLets), and action (other CmdLets)
		For AzureAD, no particular role is needed for connection and query. Get-CmdLets are available without an Admin-role.
		For MicrosoftTeams, a Teams Administrator Role is required (ideally Teams Communication or Service Administrator)
		Module MicrosoftTeams v2.0.0 now provides the CmdLets that required a Session to SkypeOnline.
    The Skype for Business Legacy Administrator Roles are still required to create the PsSession.
		Actual administrative capabilities are dependent on actual Office 365 admin role assignments (displayed as output)
		Disconnects current sessions (if found) in order to establish a clean new session to each desired service.
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Connect-Me
	.LINK
    Connect-AzureAD
	.LINK
    Connect-MicrosoftTeams
	.LINK
    Disconnect-Me
	.LINK
    Disconnect-AzureAD
	.LINK
    Disconnect-MicrosoftTeams
	#>

  [CmdletBinding()]
  [Alias('con')]
  param(
    [Parameter(Mandatory, Position = 0, HelpMessage = 'UserPrincipalName, Administrative Account')]
    [Alias('Username')]
    [string]$AccountId,

    [Parameter(HelpMessage = 'Establishes a connection to Exchange Online. Reuses credentials if authenticated already.')]
    [Alias('Exchange')]
    [switch]$ExchangeOnline,

    [Parameter(HelpMessage = 'Suppresses Session Information output')]
    [switch]$NoFeedback

  ) #param

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Required as Warnings on the OriginalRegistrarPool somehow may halt Script execution
    $WarningPreference = 'Continue'
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference') }
    if (-not $PSBoundParameters.ContainsKey('WhatIf')) { $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

    # Initialising counters for Progress bars
    [int]$step = 0
    [int]$sMax = 6

    #region Preparation
    # Preparing environment
    #Persist Stored Credentials on local machine - Value is unclear as they don't seem to be needed anymore now that New-CsOnlineSession is gone
    if (!$PSDefaultParameterValues.'Parameters:Processed') {
      $PSDefaultParameterValues.add('New-StoredCredential:Persist', 'LocalMachine')
      $PSDefaultParameterValues.add('Parameters:Processed', $true)
    }

    # Cleaning up existing sessions
    $Status = 'Preparation'
    $Operation = 'Verifying Parameters'
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"

    #region Parameter validation
    if ($PSBoundParameters.ContainsKey('ExchangeOnline')) {
      $sMax++
    }

    if ( -not $NoFeedback ) {
      $sMax = $sMax + 3
    }
    #endregion

    #Loading Modules
    $Operation = 'Loading Modules'
    $step++
    Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
    Write-Verbose -Message "$Status - $Operation"
    $AzureAdModule, $AzureAdPreviewModule, $TeamsModule, $SkypeModule = Get-NewestModule AzureAd, AzureAdPreview, MicrosoftTeams, SkypeOnlineConnector
    if ( $SkypeModule ) {
      Write-Warning -Message "Module 'SkypeOnlineConnector' detected. This module is deprecated and no longer required. If it remains on the system, it could interfere in execution of Connection Commands. Removing Module from Session - Please uninstall SkypeOnlineConnector (MSI)!"
      Remove-Module SkypeOnlineConnector -Verbose:$false -ErrorAction SilentlyContinue
    }
    Write-Verbose -Message "Importing Module 'MicrosoftTeams'"
    $SaveVerbosePreference = $global:VerbosePreference;
    $global:VerbosePreference = 'SilentlyContinue';
    Import-Module MicrosoftTeams -RequiredVersion 2.0.0 -Force -Global -Verbose:$false
    $global:VerbosePreference = $SaveVerbosePreference

    if ( $AzureAdPreviewModule ) {
      Remove-Module AzureAd -Verbose:$false -ErrorAction SilentlyContinue
      Import-Module AzureAdPreview -Force -Global -AllowClobber -Verbose:$false

      if ( -not (Assert-Module AzureAdPreview )) {
        if ( -not (Assert-Module AzureAd) ) {
          throw 'Error importing Module: Neither AzureAd nor AzureAdPreview are available'
        }
      }
    }

    # Privileged Identity Management
    # Determining options
    $Command = 'Get-AzureADMSPrivilegedRoleAssignment'
    try {
      $PIMavailable = Get-Command -Name $Command -ErrorAction Stop
      if ( $PIMavailable ) { $sMax++ }
    }
    catch {
      Write-Information "Command '$Command' not available. Privileged Identity Management role activation cannot be used. Please ensure admin roles are activated prior to running this command"
      Write-Verbose -Message 'AzureAd & MicrosoftTeams: Establishing a connection will work, though only GET-commands will be able to be executed'
      #Write-Verbose -Message "MicrosoftTeams: Executing NEW/SET/REMOVE CmdLets requires the 'Lync Administrator' ('Skype for Busines Legacy Administrator' in the Admin Center) role is not activated"
    }
    #endregion

    $ConnectionParameters = $null
    $ConnectionParameters += @{ 'ErrorAction' = 'Stop' }

    if ($PSBoundParameters.ContainsKey('Verbose')) {
      $ConnectionParameters += @{ 'Verbose' = $true }
    }
    if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
      $ConnectionParameters += @{ 'Debug' = $true }
    }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    #region Connections
    $Status = 'Establishing Connection'
    Write-Information "Establishing Connection to Tenant: $($($AccountId -split '@')[1])"
    $ConnectionOrder = @('AzureAd', 'MicrosoftTeams')
    if ($ExchangeOnline) { $ConnectionOrder += 'ExchangeOnline' }

    foreach ($Connection in $ConnectionOrder) {
      $Service = $Connection
      $step++
      $Operation = $Service
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation" -Verbose

      $MeToTheO365ServiceParams.Service = $Service
      try {
        switch ($Connection) {
          'AzureAd' {
            $AzureAdParameters = $ConnectionParameters
            $AzureAdParameters += @{ 'AccountId' = $AccountId }
            $AzureAdFeedback = Connect-AzureAD @AzureAdParameters
            #$AzureAdFeedback = Connect-MeToTheO365Service @MeToTheO365ServiceParams

            #region Activating Admin Roles
            if ( $PIMavailable ) {
              $step++
              $Operation = 'Enabling eligible Admin Roles'
              Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
              Write-Verbose -Message "$Status - $Operation" -Verbose
              try {
                $ActivatedRoles = Enable-AzureAdAdminRole -Identity $AccountId -PassThru -Force -ErrorAction Stop #(default should only enable the Teams ones? switch?)
                if ( $ActivatedRoles.Count -gt 0 ) {
                  Write-Verbose "Enable-AzureAdAdminrole - $($ActivatedRoles.Count) Roles activated." -Verbose
                }
              }
              catch {
                Write-Verbose 'Enable-AzureAdAdminrole - Tenant is not enabled for PIM' -Verbose
                $PIMavailable = $false
              }
            }
            else {
              Write-Verbose 'Enable-AzureAdAdminrole - Privileged Identity Management functions are not available'
            }
          }
          'MicrosoftTeams' {
            $MicrosoftTeamsParameters = $ConnectionParameters

            try {
              $MicrosoftTeamsFeedback = Connect-MicrosoftTeams @MicrosoftTeamsParameters
            }
            catch {
              #$MicrosoftTeamsFeedback = Connect-MeToTheO365Service @MeToTheO365ServiceParams
              $MicrosoftTeamsFeedback = Connect-MicrosoftTeams
            }
          }
          'ExchangeOnline' {
            $ExchangeOnlineParameters = $ConnectionParameters
            $ExchangeOnlineParameters += @{ 'UserPrincipalName' = $AccountId }
            $ExchangeOnlineParameters += @{ 'ShowProgress' = $true }
            $ExchangeOnlineParameters += @{ 'ShowBanner' = $false }

            $ExchangeOnlineFeedback = Connect-ExchangeOnline @ExchangeOnlineParameters
            #$ExchangeOnlineFeedback = Connect-MeToTheO365Service -AccountId $AccountId -Service $Service -ErrorAction Stop
          }
        }
      }
      catch {
        Write-Error -Message "$($_.Exception.Message)"
      }
    }

    #region Feedback
    $CsTenant = Get-CsTenant -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if ( -not $NoFeedback ) {
      $Status = 'Providing Feedback'
      $step++
      $Operation = 'Querying information about established sessions'
      Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
      Write-Verbose -Message "$Status - $Operation"

      # Preparing Output Object
      #CHECK Output Object - Write new Script `Get-ConnectMeConnection` publish as `cur`? (see profile!) Attach output to Assert with new switch?
      $SessionInfo = [PSCustomObject][ordered]@{
        Account                   = $AccountId
        AdminRoles                = $ActivatedRoles.RoleName -join ', '
        Tenant                    = ''
        TenantDomain              = ''
        TenantId                  = ''
        ConnectedTo               = [System.Collections.ArrayList]@()
        AzureEnvironment          = ''
        TeamsUpgradeEffectiveMode = ''
      }

      #AzureAd SessionInfo
      if ( Test-AzureADConnection ) {
        $SessionInfo.ConnectedTo += 'AzureAd'
        $AzureAdFeedback = Get-AzureADCurrentSessionInfo
        $SessionInfo.Tenant = "$($AccountId.split('@')[1]) - $($CsTenant.DisplayName)"
        $SessionInfo.TenantDomain = $AzureAdFeedback.TenantDomain
        $SessionInfo.TenantId = $AzureAdFeedback.TenantId
        $SessionInfo.AzureEnvironment = $AzureAdFeedback.Environment
      }
      #MicrosoftTeams SessionInfo
      if ( Test-MicrosoftTeamsConnection ) {
        $SessionInfo.ConnectedTo += 'MicrosoftTeams'
        $SessionInfo.TeamsUpgradeEffectiveMode = $CsTenant.TeamsUpgradeEffectiveMode
      }
      #Exchange SessionInfo
      if ( Test-ExchangeOnlineConnection ) {
        $SessionInfo.ConnectedTo += 'ExchangeOnline'
        #What to add?
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name): ExchangeOnlineFeedback:", ($ExchangeOnlineFeedback | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }
      }

      #Querying Admin Roles
      if ( -not $SessionInfo.AdminRoles ) {
        #AdminRoles is already populated if they have been activated with PIM (though only with eligible ones) this overwrites the previous set of roles
        $step++
        $Operation = 'Querying assigned Admin Roles'
        Write-Progress -Id 0 -Status $Status -CurrentOperation $Operation -Activity $MyInvocation.MyCommand -PercentComplete ($step / $sMax * 100)
        Write-Verbose -Message "$Status - $Operation"

        if ( Test-AzureADConnection) {
          if ( $AzureAdPreviewModule ) {
            $Roles = $(Get-AzureAdAdminRole (Get-AzureADCurrentSessionInfo).Account).RoleName -join ', '
          }
          else {
            #$Roles = $(Get-AzureAdAssignedAdminRoles (Get-AzureADCurrentSessionInfo).Account).DisplayName -join ', '
            Write-Warning -Message 'Module AzureAdPreview not present. Admin Roles cannot be enumerated.'
          }
          $SessionInfo.AdminRoles = $Roles
        }
      }

      #Output
      Write-Output $SessionInfo


      Write-Host "$(Get-Date -Format 'dd MMM yyyy HH:mm') | Ready" -ForegroundColor Green
      Get-RandomQuote
    }

    Write-Progress -Id 0 -Status 'Complete' -Activity $MyInvocation.MyCommand -Completed
    #endregion

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} # Connect-Me
