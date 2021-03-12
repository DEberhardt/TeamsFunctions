# Module:     TeamsFunctions
# Function:   UserAdmin
# Author:     David Eberhardt
# Updated:    20-DEC-2020
# Status:     Live


#TODO: Privileged Admin Groups buildout

function Enable-AzureAdAdminRole {
  <#
  .SYNOPSIS
    Enables eligible Admin Roles
  .DESCRIPTION
    Azure Ad Privileged Identity Management can require you to activate Admin Roles.
    Eligibe roles or groups can be activated with this Command
  .PARAMETER Identity
    Username of the Admin Account to enable roles for
  .PARAMETER Reason
    Optional. Small statement why these roles are requested
    By default, "admin" is used as the reason.
  .PARAMETER EnableAll
    By default, enables only Roles required to Administer Teams
    These are: Lync Administrator, User Administrator, License Administrator,
    Teams Communications Administrator, Teams Service Administrator
    This switch, when used, tries to enable all found admin role
  .PARAMETER Duration
    Optional. Integer. By default, enables Roles for 4 hours.
    Depending on your Administrators settings, values between 1 and 24 hours can be specified
  .PARAMETER TicketNr
    Optional. Integer. Only used if provided
    Depending on your Administrators settings, a ticket number may be required to process the request
  .PARAMETER Extend
    Optional. Switch. If an assignment is already active, it can be extended.
    This will leave an open request which can be closed manually.
  .PARAMETER ProviderId
    Optional. Default is 'aadRoles' for the ProviderId, however, this script could also be used for activating
    Azure Resources ('azureResources'). Use with Confirm and EnableAll.
  .PARAMETER PassThru
    Optional. Displays output object for each activated Role
    Used for further processing to verify command was successful
  .EXAMPLE
    Enable-AzureAdAdminRole John@domain.com
    Enables all eligible Teams Admin roles for User John@domain.com
  .EXAMPLE
    Enable-AzureAdAdminRole John@domain.com -EnableAll -Reason "Need to provision Users" -Duration 4
    Enables all eligible Admin roles for User John@domain.com with the reason provided.
  .EXAMPLE
    Enable-AzureAdAdminRole John@domain.com -EnableAll -ProviderId azureResources -Confirm
    Enables all eligible Azure Resources for User John@domain.com with confirmation for each Resource.
  .EXAMPLE
    Enable-AzureAdAdminRole John@domain.com -Extend -Duration 3
    If already activated, will extend the Azure Resources for User John@domain.com for up to 3 hours.
  .INPUTS
    System.String
  .OUTPUTS
    None
  .NOTES
    Limitations: MFA must be authorised first
    Currently no way to trigger it via PowerShell. If the activation fails, please sign into Office.com
    Once Authorised, this command can be used to activate your eligible Admin Roles.
    AzureResources provider activation is not yet tested.

    Thanks to Nathan O'Bryan, MVP|MCSM - nathan@mcsmlab.com for inspiring this script through Activate-PIMRole.ps1
  .COMPONENT
    UserAdmin
  .ROLE
    Activating Admin Roles
  .FUNCTIONALITY
    Enables eligible Privileged Identity roles for Administration of Teams
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Enable-AzureAdAdminRole
  .LINK
    Get-AzureAdAdminRole
  #>

  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([Void])]

  param(
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'Enter the identity of the Admin Account')]
    [Alias('UPN', 'UserPrincipalName', 'Username')]
    [string]$Identity,

    [Parameter(HelpMessage = 'Optional Reason for the request')]
    [string]$Reason,

    [Parameter(HelpMessage = 'Integer in hours to activate role(s) for')]
    [int]$Duration,

    [Parameter(HelpMessage = 'Ticket Number for use to provide to the request')]
    [int]$TicketNr,

    [Parameter(HelpMessage = 'Azure ProviderId to be used')]
    [ValidateSet('aadRoles', 'azureResources')]
    [string]$ProviderId = 'aadRoles',

    [Parameter(HelpMessage = 'Tries to extend the activation.')]
    [switch]$Extend,

    [Parameter(HelpMessage = 'Displays output of activated roles to verify')]
    [switch]$PassThru,

    [Parameter(HelpMessage = 'Overrides confirmation dialog and enables all eligible roles')]
    [switch]$Force

  ) #param

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }

    # Importing Module
    #R#equires -Modules @{ ModuleName="AzureADpreview"; ModuleVersion="2.0.2.24" }
    #TODO To be removed once AzureAd is updated containing the PIM functions and made part of the Requirements for this Module
    try {
      Write-Verbose -Message "Removing Module 'AzureAd', Importing Module 'AzureAdPreview'"
      $SaveVerbosePreference = $global:VerbosePreference;
      $global:VerbosePreference = 'SilentlyContinue';
      Remove-Module AzureAd -Force -ErrorAction SilentlyContinue -Verbose:$false
      Import-Module AzureAdPreview -Global -Force -ErrorAction Stop -Verbose:$false
      $global:VerbosePreference = $SaveVerbosePreference
    }
    catch {
      Write-Error -Message 'Module AzureAdPreview not present or failed to import. Please make sure the Module is installed'
      return
    }

    # preparing Splatting Object
    $Parameters = $null
    $Parameters += @{'ErrorAction' = 'Stop' }

    #region Supporting Parameters
    # Duration
    if ( -not $Duration ) {
      [int]$Duration = 4
      # Duration is used in $Schedule
    }

    # Reason & Ticket Number
    if ( -not $Reason ) { $Reason = 'Admin' }
    if ( $TicketNr ) { $Reason = "Ticket: $TicketNr - $Reason" }
    $Parameters += @{'Reason' = $Reason }

    # ProviderId is hardcoded (or overridden by providing a value)
    Write-Verbose -Message "Using Azure Provider Id: $ProviderId"
    $Parameters += @{'ProviderId' = $ProviderId }

    # ResourceId - is the Tenant Id
    Write-Verbose -Message 'Querying Azure Tenant Id'
    $ResourceId = (Get-AzureADCurrentSessionInfo).TenantId
    $Parameters += @{'ResourceId' = $ResourceId }

    # Assignment state is always Active - This will change for the Disable command
    $Parameters += @{'AssignmentState' = 'Active' }

    # Importing all Roles
    Write-Verbose -Message 'Querying Azure Privileged Role Definitions'
    try {
      $AllRoles = Get-AzureADMSPrivilegedRoleDefinition -ProviderId $ProviderId -ResourceId $ResourceId -ErrorAction Stop
    }
    catch {
      if ($_.Exception.Message.Contains('The tenant needs an AAD Premium 2 license')) {
        Write-Error -Message 'Cannot query role definitions. AzureAd Premium License Required' -ErrorAction Stop
      }
      else {
        Write-Error -Message "Cannot query role definitions. Exception: $($_.Exception.Message)" -ErrorAction Stop
      }
    }

    # Defining Schedule
    Write-Verbose -Message "Creating Schedule based on Duration: $Duration hours"
    $Date = Get-Date
    $start = $Date.ToUniversalTime()
    $end = $Date.AddHours($Duration).ToUniversalTime()

    $schedule = New-Object Microsoft.Open.MSGraph.Model.AzureADMSPrivilegedSchedule
    $schedule.Type = 'Once'
    $schedule.StartDateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $schedule.endDateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    Write-Verbose -Message "Admin Roles will be active for $Duration hours, until: $($end.ToString())"
    $Parameters += @{'Schedule' = $schedule }
    #endregion

    # Identity is not mandatory, using connected Session
    if ( -not $PSBoundParameters.ContainsKey('Identity') ) {
      $Identity = (Get-AzureADCurrentSessionInfo).Account.Id
      Write-Information "No Identity Provided, using user currently connected to AzureAd: '$Identity'"
    }

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    foreach ($Id in $Identity) {
      # Querying User Account
      Write-Verbose -Message "Processing User '$Id'"
      try {
        $SubjectId = (Get-AzureADUser -ObjectId "$Id" -WarningAction SilentlyContinue -ErrorAction STOP).ObjectId

        # Adding SubjectId to Parameters
        if ($Parameters.SubjectId) {
          $Parameters.SubjectId = $SubjectId
        }
        else {
          $Parameters += @{'SubjectId' = $SubjectId }
        }
      }
      catch {
        Write-Error -Message 'User Account not valid' -Category ObjectNotFound -RecommendedAction 'Verify Identity/UserPrincipalName'
        continue
      }

      # Query current Admin Roles

      <# Commented out for Admin Groups are not yet available via PowerShell
      $TenantRoles = Get-AzureADMSPrivilegedRoleAssignment -ProviderId $ProviderId -ResourceId $ResourceId #-Filter "subjectId eq '$SubjectId'"
      $MyEligibleGroups = $TenantRoles | Where-Object MemberType -EQ "Group"
      $MyRoles = $TenantRoles | Where-Object SubjectId -EQ $SubjectId
      #>
      $MyRoles = Get-AzureADMSPrivilegedRoleAssignment -ProviderId $ProviderId -ResourceId $ResourceId -Filter "subjectId eq '$SubjectId'"


      $MyActiveRoles = $MyRoles | Where-Object AssignmentState -EQ 'Active'
      $MyEligibleRoles = $MyRoles | Where-Object AssignmentState -EQ 'Eligible'
      Write-Verbose -Message "User '$Id' has currently $($MyActiveRoles.Count) of $($MyEligibleRoles.Count) activated"

      [System.Collections.ArrayList]$RolesAndGroups = @()
      <# Commented out for Admin Groups are not yet available via PowerShell
      if ($MyEligibleGroups.Count -eq 0) {
        Write-Verbose -Message "User '$Id' - No Privileged Access Groups are available that can be activated."
      #>
      if ($MyEligibleRoles.Count -eq 0) {
        if ($MyActiveRoles.Count -eq 0) {
          Write-Warning -Message "User '$Id' - No eligible Privileged Access Roles availabe!"
        }
        else {
          Write-Information "User '$Id' - No eligible Privileged Access Roles availabe, but User has $($MyActiveRoles.Count) permanently active Roles"
        }

        Continue
      }

      <# Commented out for Admin Groups are not yet available via PowerShell
      }
      else {
        # Adding Groups
        foreach ($Role in $MyEligibleGroups) {
          [void]$RolesAndGroups.Add($Role)
        }
      }
    #>
      if ($MyEligibleRoles.Count -gt 0) {
        # Adding Roles
        foreach ($Role in $MyEligibleRoles) {
          [void]$RolesAndGroups.Add($Role)
        }
      }

      # Activating Role
      [System.Collections.ArrayList]$ActivatedRoles = @()

      foreach ($R in $RolesAndGroups) {
        # Querying Role Display Name
        $RoleName = $AllRoles | Where-Object { $_.Id -eq $R.RoleDefinitionId } | Select-Object -ExpandProperty DisplayName

        # Confirm every role if not Force
        if ($PSCmdlet.ShouldProcess("$RoleName")) {
          if (-not ($Force -or $PSCmdlet.ShouldContinue("Eligible Role '$RoleName' found - Activate Role?", 'Enable-AzureAdAdminRole'))) {
            continue # user replied no
          }
          else {

            # Preparing Output object
            $ActivatedRole = @()
            $ActivatedRole = [PsCustomObject][ordered]@{
              'User'        = $Id
              'Rolename'    = $RoleName
              'Type'        = $null
              'ActiveUntil' = $null
            }

            # Adding Role Definition Id
            if ($Parameters.RoleDefinitionId) {
              $Parameters.RoleDefinitionId = $R.RoleDefinitionId
            }
            else {
              $Parameters += @{'RoleDefinitionId' = $R.RoleDefinitionId }
            }

            # Determining Activation Type (UserAdd VS UserRenew)
            <#
            NOTE The value for the Request type can be AdminAdd, UserAdd, AdminUpdate, AdminRemove, UserRemove, UserExtend, UserRenew, AdminRenew and AdminExtend.
            more options can be provided than UserExtend (Request) and UserAdd
            #>
            if ( $PSBoundParameters.ContainsKey('Extend') -and $R.RoleDefinitionId -in $MyActiveRoles.RoleDefinitionId ) {
              Write-Verbose -Message "User '$Id' - '$RoleName' is already active and will be extended"
              $ActivatedRole.Type = 'UserExtend'
              if ($Parameters.Type) {
                $Parameters.Type = 'UserExtend'
              }
              else {
                $Parameters += @{'Type' = 'UserExtend' }
              }
            }
            else {
              Write-Verbose -Message "User '$Id' - '$RoleName' is currently not active and will be activated"
              $ActivatedRole.Type = 'UserAdd'
              if ($Parameters.Type) {
                $Parameters.Type = 'UserAdd'
              }
              else {
                $Parameters += @{'Type' = 'UserAdd' }
              }
            }

            #Activating the Role
            if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
              "Function: $($MyInvocation.MyCommand.Name) - Parameters for Open-AzureADMSPrivilegedRoleAssignmentRequest", ( $Parameters | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
            }

            try {
              Write-Verbose -Message "User '$Id' - '$RoleName' - Activating Role"
              $ActivatedRole.ActiveUntil = $schedule.endDateTime
              $null = Open-AzureADMSPrivilegedRoleAssignmentRequest @Parameters
              [void]$ActivatedRoles.Add($ActivatedRole)
            }
            catch {
              if ($_.Exception.Message.Contains('ExpirationRule')) {
                # Amending Schedule
                if ($Duration -eq 4) { $Duration = 2 } else { $Duration = 4 }
                Write-Warning -Message "Specified Duration is not allowed, re-trying with $Duration hours"
                $end = $Date.AddHours($Duration).ToUniversalTime()

                $schedule.endDateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                Write-Verbose -Message "Admin Roles will be active for $Duration hours, until: $($end.ToString())"
                $Parameters.Schedule = $schedule

                try {
                  Write-Verbose -Message "User '$Id' - '$RoleName' - Activating Role"
                  $ActivatedRole.ActiveUntil = $schedule.endDateTime
                  $null = Open-AzureADMSPrivilegedRoleAssignmentRequest @Parameters
                  [void]$ActivatedRoles.Add($ActivatedRole)
                }
                catch {
                  if ($_.Exception.Message.Contains('ExpirationRule')) {
                    Write-Error -Message 'Specified Duration is not allowed, please try again with a lower number.' -Category InvalidData
                  }
                  else {
                    Write-Error -Message $_.Exception.Message
                  }
                }
              }
              else {
                Write-Error -Message $_.Exception.Message
              }

            }
          }
        }
      }

    }

    # Re-Query and output (for all Users!)
    if ( $PassThru ) {
      return $ActivatedRoles
    }

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #Enable-AzureAdAdminRole
