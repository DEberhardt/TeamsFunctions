﻿# Module:     TeamsFunctions
# Function:   AzureAd Licensing
# Author:     Jeff Brown
# Updated:    17-APR-2020
# Status:     Archived




function GetActionOutputObject2 {
  <#
  .SYNOPSIS
  Tests whether a valid PS Session exists for SkypeOnline (Teams)
  .DESCRIPTION
  Helper function for Output with 2 Parameters
  .PARAMETER Name
  Name of account being modified
  .PARAMETER Result
  Result of action being performed
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param(
    [Parameter(Mandatory = $true, HelpMessage = 'Name of account being modified')]
    [string]$Name,

    [Parameter(Mandatory = $true, HelpMessage = 'Result of action being performed')]
    [string]$Result
  )

  begin {
    Show-FunctionStatus -Level Archived
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    $outputReturn = [PSCustomObject][ordered]@{
      User   = $Name
      Result = $Result
    }

    return $outputReturn
  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end

} # GetActionOutputObject2
