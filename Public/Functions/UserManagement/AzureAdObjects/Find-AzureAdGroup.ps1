﻿# Module:     TeamsFunctions
# Function:   Lookup
# Author:     David Eberhardt
# Updated:    24-JAN-2021
# Status:     Live




function Find-AzureAdGroup {
  <#
	.SYNOPSIS
		Returns an Object if an AzureAd Group has been found
	.DESCRIPTION
		Simple lookup - does the Group Object exist - to avoid TRY/CATCH statements for processing
	.PARAMETER Identity
    Mandatory. String to search. Provide part or full DisplayName, MailAddress or MailNickName
    Returns all matching groups
	.EXAMPLE
    Find-AzureAdGroup [-Identity] "My Group"
    Will return all Groups that have "My Group" in the DisplayName, ObjectId or MailNickName
	.EXAMPLE
		Find-AzureAdGroup -Identity "MyGroup@domain.com"
    Will return all Groups that match "MyGroup@domain.com" in the DisplayName, ObjectId or MailNickName
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Find-AzureAdUser
  .LINK
    Get-AzureAdGroup
  .LINK
    Get-TeamsCallableEntity
  #>

  [CmdletBinding()]
  [OutputType([System.Object])]
  param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, HelpMessage = 'This is the Name or MailAddress of the Group')]
    [Alias('GroupName', 'Name')]
    [string]$Identity
  ) #param

  begin {
    Show-FunctionStatus -Level Live
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Adding Types
    Add-Type -AssemblyName Microsoft.Open.AzureAD16.Graph.Client
    Add-Type -AssemblyName Microsoft.Open.Azure.AD.CommonLibrary

    # Loading all Groups
    if ( -not $global:TeamsFunctionsTenantAzureAdGroups) {
      Write-Verbose -Message 'Groups not loaded yet, depending on the size of the Tenant, this will run for a while!' -Verbose
      $global:TeamsFunctionsTenantAzureAdGroups = Get-AzureADGroup -All $true -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    }

    $Groups = $null

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    [System.Collections.ArrayList]$Groups = @()

    $Groups += $global:TeamsFunctionsTenantAzureAdGroups | Where-Object DisplayName -Like "*$Identity*"
    $Groups += $global:TeamsFunctionsTenantAzureAdGroups | Where-Object Description -Like "*$Identity*"
    $Groups += $global:TeamsFunctionsTenantAzureAdGroups | Where-Object ObjectId -Like "*$Identity*"
    $Groups += $global:TeamsFunctionsTenantAzureAdGroups | Where-Object Mail -Like "*$Identity*"

    $MailNickName = $Identity.Split('@')[0]
    $Groups += $global:TeamsFunctionsTenantAzureAdGroups | Where-Object Mailnickname -Like "*$MailNickName*"

    # Output - Filtering objects
    if ( $Groups ) {
      $Groups | Sort-Object -Unique -Property ObjectId | Get-Unique
    }
  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} # Find-AzureAdGroup