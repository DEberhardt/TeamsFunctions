﻿# Module:   TeamsFunctions
# Function: Testing
# Author:		David Eberhardtt
# Updated:  01-JUN-2020
# Status:   PreLive

function Test-ExchangeOnlineConnection {
  <#
	.SYNOPSIS
		Tests whether a valid PS Session exists for ExchangeOnline
	.DESCRIPTION
		A connection established via Connect-ExchangeOnline is parsed.
		This connection must be valid (Available and Opened)
	.EXAMPLE
		Test-ExchangeOnlineConnection
		Will Return $TRUE only if a session is found.
  #>

  [CmdletBinding()]
  [OutputType([Boolean])]
  param() #param

  $Sessions = Get-PSSession -WarningAction SilentlyContinue
  if ([bool]($Sessions.Computername -match "outlook.office365.com")) {
    $PSExchangeOnlineSession = $Sessions | Where-Object { $_.State -eq "Opened" -and $_.Availability -eq "Available" }
    if ($PSExchangeOnlineSession.Count -ge 1) {
      return $true
    }
    else {
      return $false
    }
  }
  else {
    return $false
  }

} #Test-ExchangeOnlineConnection