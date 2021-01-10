﻿# Module:     TeamsFunctions
# Function:   Assertion
# Author:     David Eberhardt
# Updated:    15-DEC-2020
# Status:     Live




function Show-FunctionStatus {
  <#
	.SYNOPSIS
		Gives Feedback of FunctionStatus
	.DESCRIPTION
    On-Screen Output depends on Parameter Level
  .PARAMETER Level
    Level of Detail
	.EXAMPLE
    Show-FunctionStatus -Level Deprecated
    Indicates that the Function is deprecated.
  .NOTES
    This will only ever show the status of the first Command in the Stack (i.E. when called from a function).
    It will not display the same information for any nested commands.
    Available options are:
    Alpha:      Function in development. No guarantee of functionality. Here be dragons.
    Beta:       Function in development. No guarantee of functionality
    RC:         Release Candidate. Functionality is built
    Prelive:    Live function that is only lacking Pester tests.
    Live:       Live function that has proven with tests or without that it delivers.

    Unmanaged:  Legacy Function from SkypeFunctions, not managed
    Deprecated: Function flagged for removal/replacement
    Archived:   Function is archived
  #>

  [CmdletBinding()]
  param(
    [Validateset("Alpha", "Beta", "RC", "PreLive", "Live", "Unmanaged", "Deprecated", "Archived")]
    $Level
  ) #param

  $Stack = Get-PSCallStack
  if ($stack.length -gt 3) {
    return
  }
  else {
    $Function = ($Stack | Select-Object -First 2).Command[1]

    switch ($Level) {
      "Alpha" {
        $DebugPreference = "Inquire"
        $VerbosePreference = "Continue"
        Write-Debug -Message "$Function has [ALPHA] Status: It may not work as intended or contain serious gaps in functionality. Please handle with care" -Debug
      }
      "Beta" {
        $DebugPreference = "Continue"
        $VerbosePreference = "Continue"
        Write-Debug -Message "$Function has [BETA] Status: Build is not completed, functionality missing or parts untested. Please report issues via GitHub"
      }
      "RC" {
          Write-Verbose -Message "$Function has [RC] Status: Functional, but still being tested. Please report issues via GitHub" -Verbose
      }
      "PreLive" {
        Write-Verbose -Message "$Function has [PreLIVE] Status. Should you encounter issues, please get in touch via GitHub or 'TeamsFunctions@outlook.com'"
      }
      "Live" {
        Write-Verbose -Message "$Function is [LIVE]. Should you encounter issues, please get in touch via GitHub or 'TeamsFunctions@outlook.com'"
      }
      "Unmanaged" {
        Write-Verbose -Message "$Function is [LIVE] but [UNMANAGED] and comes as-is."
      }
      "Deprecated" {
        Write-Verbose -Message "$Function is [LIVE] but [DEPRECATED]!" -Verbose
      }
      "Archived" {
        Write-Verbose -Message "$Function is [ARCHIVED]!" -Verbose
      }
    }

  }
} #Show-FunctionStatus
