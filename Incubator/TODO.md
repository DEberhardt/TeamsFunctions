# ToDo List

## Documentation

- [x] TeamsFunctions-help.xml (now in DOCS!)
- [x] Update Markdown files with Platypus does not automatically update .MD files...

- [x] Create Docs
- [x] Create Help
- [ ] Populate about-help
  - [x] Write Licensing
  - [x] Write SupportFunctions
  - [x] Write AutoAttendant
  - [x] Write CallableEntities
  - [x] Write CallQueue
  - [x] Write CommonAreaPhone
  - [ ] Write AnalogDevices
  - [x] Write ResourceAccount
  - [x] Write Session
  - [x] Write UserManagement
  - [x] Write VoiceConfiguration
- [ ] Add Examples to About
  - [x] Examples for Licensing
  - [x] Examples for SupportFunctions
  - [x] Examples for CallableEntities
  - [x] Examples for CallQueue
  - [x] Examples for Session
  - [x] Examples for UserManagement
  - [x] Examples for VoiceConfiguration
  - [ ] Example for Voice Route(!)

## Module MicrosoftTeams

- [x] New-CsOnlineSession does not have -Username anymore. Test MFA with Credential instead (also interoperability with Connect-Me/AzureAD and MicrosoftTeams)
- [x] Rework Connect-SkypeOnline to use MicrosoftTeams
- [x] Evaluate missing SSO (Authentication once, selection of Account still there)
- [x] Remove SkypeOnlineConnector

## TeamsAutoAttendant

- [ ] Add features
  - [x] Operator
  - [x] Call target
  - [x] Default Schedules
  - [ ] Default HolidaySets for a Country (this year/next year)
- [ ] Evaluate integration Scripts
  - [ ] Default forward to CQ (wrapper for New-AA, with Menu to Forward to CallTarget Queue)
  - [ ] Construct of 1 AA and 1 CQ, 2 RA, 2 RAA?
- [ ] Continue to test Scripts

## Licenses

- [x] Figure out how to enable disabled plans
- [ ] Evaluate Enable/Disable CmdLet rather (or in addition to) Set-AzureAdUserLicenseServicePlan
- [ ] TBC

## Code Improvements

- [x] Create Function Template
- [x] Add Help
- [ ] Pipeline
  - [x] Add Pipeline options
  - [x] Pipeline active by default (design for Pipes)
  - [ ] Test Pipelines further
- [ ] Disable Positional Binding for CQ and AA scripts https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute?view=powershell-5.1
- [ ] Add Supports Paging (First, Last) for Get Commands Get-TeamsCallQueue and Get-TeamsAutoAttendant at least!
- [ ] Add Timestamp to Verbose steps when Processing multiple elements (just inside the ForEach)
- [ ] Add Argument Completer to some functions, where appropriate: [ArgumentCompleter({(Get-Eventlog -List).log})]
- [ ] Test for PowerShell 7
- [ ] TBC

## Pester

- [ ] Change Module checks to this format: https://vexx32.github.io/2020/07/08/Verify-Module-Help-Pester/
- [ ] Figure out Testing with active Sessions - Do I need a demo Tenant?
- [ ] Add more tests
- [ ] MORE

## Other things

### Evaluate Export to CLIXML

```powershell
$i = Import-Clixml -Path "XML FILE"
$i = Import-Clixml -Path E:\File.xml
$commandresult = $i['Get-CsOnlineUser']
$commandresult | Where-Object SipAddress -eq _____sip:xxx
```
