﻿# Module:   TeamsFunctions
# Function: Testing
# Author:   David Eberhardt
# Updated:  01-SEP-2020
# Status:   Live




function Test-SkypeOnlineConnection {
  <#
  .SYNOPSIS
    Tests whether a valid PS Session exists for SkypeOnline (Teams)
  .DESCRIPTION
    A connection established via Connect-SkypeOnline is parsed.
    This connection must be valid (Available and Opened)
  .EXAMPLE
    Test-SkypeOnlineConnection
    Will Return $TRUE only if a valid and open session is found.
  .INPUTS
    System.Void
  .OUTPUTS
    System.Boolean
  .NOTES
    Added check for Open Session to err on the side of caution.
    Use with Disconnect-SkypeOnline when tested negative, then Connect-SkypeOnline
  .COMPONENT
    TeamsSession
  .FUNCTIONALITY
    Tests the connection to MicrosoftTeams (SkypeOnline)
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Test-SkypeOnlineConnection.md
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_TeamsSession.md
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  #>

  [CmdletBinding()]
  [OutputType([Boolean])]
  param() #param

  begin {
    Show-FunctionStatus -Level Deprecated
    #Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"

  } #begin

  process {
    #Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    $Sessions = Get-PSSession -WarningAction SilentlyContinue
    $Sessions = $Sessions | Where-Object { $_.Computername -match 'online.lync.com' -or $_.ComputerName -eq 'api.interfaces.records.teams.microsoft.com' }
    if ($Sessions.Count -ge 1) {
      #Write-Verbose "Teams Session found"
      $Sessions = $Sessions | Where-Object { $_.State -eq 'Opened' -and $_.Availability -eq 'Available' }
      if ($Sessions.Count -ge 1) {
        #Write-Verbose "Teams Session found, open and valid"
        return $true
      }
      else {
        return $false
      }
    }
    else {
      return $false
    }

  } #process

  end {
    #Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end

} #Test-SkypeOnlineConnection
