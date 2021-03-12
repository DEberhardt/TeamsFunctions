# Teams Session CmdLets

## about_TeamsSession

## SHORT DESCRIPTION

Connecting to the Teams Backend is, for now, still done through SkypeOnline

## LONG DESCRIPTION

SkypeOnline and MSOnline (AzureADv1) are the two oldest Office 365 Services. Creating a Session to them is not implemented very nicely. These CmdLets try to address this.

The introduction of Privileged Identity Management and Privileged Access Groups further requires some manual steps that these are trying to make simpler and provide an easier way to connect and activate your roles

## CmdLets

| Function                                                    | Description                                                                                                                                  |
| -----------------------------------------------------------: | -------------------------------------------------------------------------------------------------------------------------------------------- |
| [`Connect-SkypeOnline`](../docs/Connect-SkypeOnline.md)       | Creates a Session to SkypeOnline                                                                            |
| [`Connect-Me`](../docs/Connect-Me.md) (con)                   | Creates a Session to SkypeOnline and AzureAD in one go. Only displays **ONE** authentication prompt, and, if applicable, **ONE** MFA prompt! Also tries to enable your Admin Roles in PIM. |
| [`Disconnect-SkypeOnline`](../docs/Disconnect-SkypeOnline.md) | Disconnects from a Session to SkypeOnline. This helps preventing timeouts and hanging sessions                                                       |
| [`Disconnect-Me`](../docs/Disconnect-Me.md) (dis)             | Disconnects form all Sessions to SkypeOnline, MicrosoftTeams and AzureAD                                                                     |

Connect-Me aims to solve the issue of having and maintaining a connection to all Office 365 needed for Administration. In most cases, that will be AzureAd and SkypeOnline, but could also need ExchangeOnline and MicrosoftTeams (which is now being tethered to SkypeOnline if the CsOnlineSession is created with the Module MicrosoftTeams)

## Admin Roles

Activating Admin Roles made easier. Please note that Privileged Access Groups are not yet integrated as there are no PowerShell commands available yet in the AzureAdPreview Module. This will be added as soon as possible. Commands are used with `Connect-Me`, but can be used on its own just as well.

> [!NOTE] Please **note**, that Privileged Admin Groups are currently not covered by these CmdLets. This will be added as soon as they have been fully documented and PowerShell CmdLets are available for them.

| Function                                                      | Description                                                                                                                                     |
| -------------------------------------------------------------: | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| [`Enable-AzureAdAdminRole`](../docs/Enable-AzureAdAdminRole.md) | Enables Admin Roles assigned directly to the AccountId provided. If no accountId is provided, the currently connected User to AzureAd is taken. |
| [`Get-AzureAdAdminRole`](../docs/Get-AzureAdAdminRole.md)       | Displays all (active or eligible) Admin Roles assigned to an AzureAdUser                                                                        |

### Support Functions

| Function                                                                      | Description                                                                                   |
| -----------------------------------------------------------------------------: | --------------------------------------------------------------------------------------------- |
| [`Assert-AzureAdConnection`](../docs/Assert-AzureAdConnection.md)               | Tests connection and visual feedback in the Verbose stream if called directly.                |
| [`Assert-MicrosoftTeamsConnection`](../docs/Assert-MicrosoftTeamsConnection.md) | Tests connection and **Attempts to reconnect** a timed-out session. Alias `PoL` *Ping-of-life*                |
| [`Test-AzureAdConnection`](../docs/Test-AzureAdConnection.md)                   | Verifying a Session to AzureAD exists                                                         |
| [`Test-SkypeOnlineConnection`](../docs/Test-SkypeOnlineConnection.md)           | Verifying a Session to SkypeOnline exists                                                     |
| [`Test-ExchangeOnlineConnection`](../docs/Test-ExchangeOnlineConnection.md)     | Verifying a Session to ExchangeOnline exists                                                  |

The Assert cmdlets are nested in all Scripts to ensure sessions are created and available

## EXAMPLES

### Example 1 - Connecting to AzureAd, MicrosoftTeams

````powershell
Connect-Me [-AccountId] John@domain.com
# Establishes a session to AzureAd, enables Admin Roles (with AzureAdPreview), Connects to MicrosoftTeams
````

### Example 2 - Connecting to AzureAd, MicrosoftTeams and ExchangeOnline

````powershell
Connect-Me [-AccountId] John@domain.com -Exchange
# Establishes a session to AzureAd, enables Admin Roles (with AzureAdPreview), Connects to MicrosoftTeams and ExchangeOnline
````

## NOTE

To properly administer Teams, a connection to `AzureAd` is most likely needed. Privileged Identity Management and Role Activation are currently only available with the Module `AzureAdPreview` installed in Version `2.0.2.24` or higher, until the functions become generally available through the AzureAd Module.

> [!NOTE] Most Functions and CmdLets in this module rely on a connection to AzureAd as well as SkypeOnline.

Initially, this module was built around the use of the Module `SkypeOnlineConnector`(v7). The Connector is now deprectated and this Module does no longer require it to be installed.
The command to establish a connection has been ported to the Module `MicrosoftTeams` (in v1.1.6) which is a requirement for this Module.

The Connection CmdLet does behave slightly differently: Seamless Single-Sign-On is not available, the Account needs to be selected manually. Please see detailed notes in [`Connect-Me`](../docs/Connect-Me.md) and [`Connect-SkypeOnline`](../docs/Connect-SkypeOnline.md) for details.

## Development Status

Mature, but a bit of a moving target - Continuously 'in Progress'. An update to the MicrosoftTeams Command and documentation is anticipated that would enable Single-sign-on with MFA enabled. Issues have been raised via Github to address.

## TROUBLESHOOTING NOTE

Thoroughly tested, but Unit-tests for these CmdLets are not yet available.

Please disconnect your session cleanly, before reconnecting. I have found that, when using the MicrosoftTeams Module, the whole PowerShell Session needs to be recreated which is annoying. This is still evaluated and tested with every PreRelease.

The Session can also currently not be reconnected to, a credential Dialog is shown with the Username "oAuth". Please cancel this and the following empty credential dialog and run Connect-Me again. This will try to reconnect. I am trying to write a script that can catch that and try to reconnect you more cleanly.

## SEE ALSO

- [`Connect-AzureAd`](https://docs.microsoft.com/en-us/powershell/module/azureAd/connect-azuread)
- [`Connect-microsoftteams`](https://docs.microsoft.com/en-us/powershell/module/teams/connect-microsoftteams?view=teams-ps)

## KEYWORDS

- SkypeOnline
- MicrosoftTeams
- Module
- Requirements
