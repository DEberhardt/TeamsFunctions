﻿# Module:     TeamsFunctions
# Function:   Lookup
# Author:     David Eberhardt
# Updated:    14-NOV-2020
# Status:     PreLive

function Find-AzureAdUser {
  <#
	.SYNOPSIS
		Returns User Object in Azure AD from a provided UPN
	.DESCRIPTION
    Enables UPN lookup for AzureAD users
    This simplifies the query without having to rely on -ObjectId or -SearchString parameters in Get-AzureAdUser
	.PARAMETER Identity
		Required. The sign-in address or User Principal Name of the user account to query.
	.EXAMPLE
		Find-AzureAdUser John@domain.com
		Will Return the Azure AD Object for John@domain.com, otherwise returns error message from Get-AzureAdUser
  .INPUTS
    System.String
  .OUTPUTS
    Microsoft.Open.AzureAD.Model.User
	#>
  [CmdletBinding()]
  [OutputType([Microsoft.Open.AzureAD.Model.User])]
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "This is the UserID (UPN)")]
    [Alias('UserPrincipalName')]
    [string[]]$Identity
  ) #param

  begin {
    Show-FunctionStatus -Level PreLive
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Adding Types
    Add-Type -AssemblyName Microsoft.Open.AzureAD16.Graph.Client
    Add-Type -AssemblyName Microsoft.Open.Azure.AD.CommonLibrary
  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"
    foreach ($UPN in $Identity) {
      try {
        # This is functional but slow in bigger environments!
        #$User = Get-AzureADUser -All:$true | Where-Object {$_.UserPrincipalName -eq $UPN} -ErrorAction STOP
        $User = Get-AzureADUser -ObjectId "$UPN" -WarningAction SilentlyContinue -ErrorAction STOP
        Write-Output $User
      }
      catch [Microsoft.Open.AzureAD16.Client.ApiException] {
        Write-Verbose -Message "User '$UPN' not found" -Verbose
        return $null
      }
      catch {
        Write-Verbose -Message "User '$UPN' not found" -Verbose
        return $null
      }
    }

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #Find-AzureAdUser