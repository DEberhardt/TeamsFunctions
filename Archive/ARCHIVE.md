# Teams Functions - Archive

Call it the cutting room floor, an archive or just a repository of unpublished functions.
They will remain in the GitHub repository in the ARCHIVE folder, but will not be published or used in the main module.

Here are some of these functions:

| Function                        | Description                                                                                    | Replacement                       |
| ------------------------------- | ---------------------------------------------------------------------------------------------- | --------------------------------- |
| `Add-TeamsUserLicense`          | Adds one or more Licenses specified per Switch to the provided Identity                        | `Set-TeamsUserLicense`            |
| `Assert-SkypeOnlineConnection`  | Tests connection and **Attempts to reconnect** a timed-out session. Alias `PoL` _Ping-of-life_ | `Assert-MicrosoftTeamsConnection` |
| `Get-AzureAdAssignedAdminRoles` | v1 of the Admin Role query Function                                                            | `Get-AzureAdAdminRole`            |
| `Get-SkuIdFromSkuPartNumber`    | Helper function for Licensing. Returns a SkuID from a specific SkuPartNumber                   | `Get-AzureAdLicense`              |
| `Get-SkuPartNumberFromSkuId`    | Helper function for Licensing. Returns a SkuPartNumber from a specific SkuID                   | `Get-AzureAdLicense`              |
| `Get-TeamsLicenseServicePlan`   | v1 of the Licensing Function for Service Plans                                                 | `Get-AzureAdLicenseServicePlan`   |
| `Get-TeamsLicense`              | v1 of the Licensing Function for Licenses                                                      | `Get-AzureAdLicense`              |
| `GetActionOutputObject2`        | Private Function and like Write-ErrorRecord a way to display output                            |                                   |
| `GetActionOutputObject3`        | Private Function and like Write-ErrorRecord a way to display output                            |                                   |
| `ProcessLicense`                | Private Function and the gears behind `Add-TeamsUserLicense`                                   |                                   |
| `Set-TeamsUserPolicy`           | Assigns specific Policies to a User. Currently only six policies available                     | None                              |
| `Test-SkypeOnlineConnection`    | Verifying a Session to SkypeOnline exists                                                      | `Test-MicrosoftTeamsConnection`   |
| `Test-TeamsTenantPolicy`        | Tests whether any Policy is present in the Tenant. Used Invoke-Expression                      | None                              |
| `Write-ErrorRecord`             | Troubleshooting and to display Errors in a more readable format and in the Output stream       | `Write-Error`                     |
