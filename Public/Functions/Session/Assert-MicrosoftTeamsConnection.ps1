# Module:     TeamsFunctions
# Function:   Session
# Author:     David Eberhardt
# Updated:    01-OCT-2020
# Status:     Live




function Assert-MicrosoftTeamsConnection {
  <#
	.SYNOPSIS
		Asserts an established Connection to MicrosoftTeams
	.DESCRIPTION
		Tests a connection to MicrosoftTeams is established.
	.EXAMPLE
		Assert-MicrosoftTeamsConnection
    Will run Test-MicrosoftTeamsConnection and, if successful, stops.
    If unsuccessful, displays request to create a new session and stops.
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  #>

  [CmdletBinding()]
  [Alias('pol')]
  [OutputType([Boolean])]
  param() #param

  begin {
    Show-FunctionStatus -Level Live
    $Stack = Get-PSCallStack
    $Called = ($stack.length -ge 3)

  } #begin

  process {

    if (Test-MicrosoftTeamsConnection) {
      if ($stack.length -lt 3) {
        Write-Verbose -Message '[ASSERT] MicrosoftTeams Session Connected'
      }
      return $(if ($Called) { $true })
    }
    else {
      try {
        $null = Connect-MicrosoftTeams -ErrorAction Stop
        Start-Sleep -Seconds 2
        if (Test-MicrosoftTeamsConnection) {
          if ($stack.length -lt 3) {
            Write-Verbose -Message '[ASSERT] MicrosoftTeams Session Reconnected!'
          }
          return $(if ($Called) { $true })
        }
        else {
          throw "Reconnect failed"
        }
      }
      catch {
        if ($PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -eq 'Continue') {
          "Function: $($MyInvocation.MyCommand.Name) - Exception message", ( $_.Exception.Message | Format-Table -AutoSize | Out-String).Trim() | Write-Debug
        }
        if (Test-AzureADConnection) {
          $AzureAdFeedback = Get-AzureADCurrentSessionInfo
          Write-Host '[ASSERT] ERROR: Reconnect unsuccessful. Trying to disconnect and reconnect you. Please validate your Admin roles, disconnect and reconnect' -ForegroundColor Red
          try {
            Disconnect-Me
            Connect-Me -AccountId $($AzureAdFeedback.Account) -NoFeedback
            return $(if ($Called) { $true })
          }
          catch {
            return $(if ($Called) { $false })
          }
        }
        else {
          Write-Host '[ASSERT] ERROR: Reconnect unsuccessful. Connect-MicrosoftTeams failed and no Session to AzureAd exists. Please validate your Admin roles, disconnect and reconnect' -ForegroundColor Red
          return $(if ($Called) { $false })
        }
        <# Commented out to avoid having two authentication popups
        else {
          $null = Connect-MicrosoftTeams -ErrorAction SilentlyContinue
        }
        #>
      }
    }
  } #process

  end {

  } #end
} #Assert-MicrosoftTeamsConnection
