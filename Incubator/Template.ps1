# Module:   TeamsFunctions
# Function:
# Author:	  David Eberhardt
# Updated:  xx-xxxx-2021
# Status:   Alpha




function Verb-Noun {
  <#
  .SYNOPSIS
    Short description
  .DESCRIPTION
    Long description
  .PARAMETER Identity
    x
  .PARAMETER x
    x
  .EXAMPLE
    Verb-Noun -Identity John@domain.com
    xx
  .INPUTS
    System.String
  .OUTPUTS
    System.Object
  .NOTES
    xx
  .COMPONENT
    xx
  .ROLE
    xx
  .FUNCTIONALITY
    xx
  .LINK
    https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/
  .LINK
    Verb2-Noun
  #>

  [CmdletBinding()]
  [Alias('')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, HelpMessage = 'Username(s)')]
    [Alias('', '')]
    [string[]]$Identity,

    [Parameter(HelpMessage = '')]
    [Alias('')]
    [String]$x

  )

  begin {
    Show-FunctionStatus -Level Alpha
    Write-Verbose -Message "[BEGIN  ] $($MyInvocation.MyCommand)"
    Write-Verbose -Message "Need help? Online:  $global:TeamsFunctionsHelpURLBase$($MyInvocation.MyCommand)`.md"

    # Asserting AzureAD Connection
    if (-not (Assert-AzureADConnection)) { break }

    # Asserting SkypeOnline Connection
    if (-not (Assert-SkypeOnlineConnection)) { break }

    # Setting Preference Variables according to Upstream settings
    if (-not $PSBoundParameters.ContainsKey('Verbose')) { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
    if (-not $PSBoundParameters.ContainsKey('Confirm')) { $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference') }
    if (-not $PSBoundParameters.ContainsKey('WhatIf')) { $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference') }
    if (-not $PSBoundParameters.ContainsKey('Debug')) { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') } else { $DebugPreference = 'Continue' }
    if ( $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = $PSCmdlet.SessionState.PSVariable.GetValue('InformationAction') } else { $InformationPreference = 'Continue' }


  } #begin

  process {
    Write-Verbose -Message "[PROCESS] $($MyInvocation.MyCommand)"

    foreach ($Id in $Identity) {
      Write-Verbose -Message "[PROCESS] Processing '$Id'"


      #TODO Todo-Tree Colour Scheme test:

      #BUG    White on red, full line, Bug icon in gutter
      #TODO   White on GREEN, text only, beaker, gutter
      #BODGE  White on MAGENTA, text only, TBD, gutter
      #HACK   same as bodge
      #FIXME  BLACK on Orange, text only, FIRE (default), gutter
      #FIXIT  Same as Fixme
      #FIX    Same as Fixme
      #CHECK  White on BLUE, text only, TBD, gutter
      #TEST   White on CYAN, text only, TBD, gutter
      #VALIDATE   Same as Test
      #IMPROVE  Check performance
      #DOC    White on GREY?, text only, SHEET?
      #WRITE  Same as DOC
      #NOTE   Same as DOC?
      [ ] Open task
      [x] Closed task

    } #foreach Identity

  } #process

  end {
    Write-Verbose -Message "[END    ] $($MyInvocation.MyCommand)"
  } #end
} #Verb-Noun