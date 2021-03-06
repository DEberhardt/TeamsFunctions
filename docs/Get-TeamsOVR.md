---
external help file: TeamsFunctions-help.xml
Module Name: TeamsFunctions
online version: https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-TeamsOVR.md
schema: 2.0.0
---

# Get-TeamsOVR

## SYNOPSIS
Lists all Online Voice Routes by Name

## SYNTAX

```
Get-TeamsOVR [[-Identity] <String>] [<CommonParameters>]
```

## DESCRIPTION
To quickly find Online Voice Routes to troubleshoot, an Alias-Function to Get-CsOnlineVoiceRoute

## EXAMPLES

### EXAMPLE 1
```
Get-TeamsOVR
```

Returns the Object for all Online Voice Routes (except "LocalRoute")
Behaviour like: Get-CsOnlineVoiceRoute, if more than 3 results are found, only names are returned

### EXAMPLE 2
```
Get-TeamsOVR -Identity OVR-EMEA-National
```

Returns the Object for the Online Voice Route "OVR-EMEA-National" (provided it exists).
Behaviour like: Get-CsOnlineVoiceRoute -Identity "OVR-EMEA-National"

### EXAMPLE 3
```
Get-TeamsOVR -Identity OVR-EMEA-*
```

Lists Online Voice Routes with "OVR-EMEA-" in the Name
Behaviour like: Get-CsOnlineVoiceRoute -Filter "*OVR-EMEA-*"

## PARAMETERS

### -Identity
String.
Name or part of the Voice Route.
Can be omitted to list Names of all Routes (except "Global").
If provided without a '*' in the name, an exact match is sought.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None
### System.String
## OUTPUTS

### System.Object
## NOTES
This script is indulging the lazy admin.
It behaves like Get-CsOnlineVoiceRoute with a twist:
If more than three results are found, a reduced set of Parameters are shown for better visibility:
Get-CsOnlineVoiceRoute | Where-Object Identity -NE 'LocalRoute' | Select-Object Identity, Priority, NumberPattern, OnlinePstnGatewayList

## RELATED LINKS

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-TeamsOVR.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-TeamsOVR.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_VoiceConfiguration.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_VoiceConfiguration.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_Supporting_Functions.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_Supporting_Functions.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/)

