# Module:   TeamsFunctions
# Function: Lookup
# Author:	  David Eberhardt
# Updated:  10-JAN-2021
# Status:   PreLive




function Remove-TeamsFunctionsGlobalVariable {
  <#
  .SYNOPSIS
    Removes Global Variables set by CmdLets in TeamsFunctions
  .DESCRIPTION
    Global Variables are set during the session which may include Tenant Data
    These Variables are removed with this Command
  .EXAMPLE
    Remove-TeamsFunctionsGlobalVariable
    Removes all Global Variables set by CmdLets in TeamsFunctions
  .INPUTS
    System.Void
  .OUTPUTS
    System.Void
  .FUNCTIONALITY
    Helper Function
  #>

  param ()

  $VariableNames = @(
    "TeamsFunctionsMSTelephoneNumbers", # Used for Microsoft TelephoneNumbers from the Tenant
    "TeamsFunctionsMSAzureAdLicenses", # Used for all Licensing commands
    "TeamsFunctionsMSAzureAdLicenseServicePlans" # Used for all Licensing commands
  )

  $null = (Remove-Variable -Name $VariableNames -Scope Global -ErrorAction SilentlyContinue)

} #Remove-TeamsFunctionsGlobalVariable
