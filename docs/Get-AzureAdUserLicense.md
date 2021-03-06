---
external help file: TeamsFunctions-help.xml
Module Name: TeamsFunctions
online version: https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-AzureAdUserLicense.md
schema: 2.0.0
---

# Get-AzureAdUserLicense

## SYNOPSIS
Returns License information for an Object in AzureAD

## SYNTAX

```
Get-AzureAdUserLicense [-UserPrincipalName] <String[]> [-FilterRelevantForTeams] [<CommonParameters>]
```

## DESCRIPTION
Returns an Object containing all Licenses found for a specific Object
Licenses and ServicePlans are nested in the respective parameters for further investigation

## EXAMPLES

### EXAMPLE 1
```
Get-AzureAdUserLicense [-UserPrincipalname] John@domain.com
```

Displays all licenses assigned to User John@domain.com

### EXAMPLE 2
```
Get-AzureAdUserLicense -UserPrincipalname John@domain.com,Jane@domain.com
```

Displays all licenses assigned to Users John@domain.com and Jane@domain.com

### EXAMPLE 3
```
Get-AzureAdUserLicense -UserPrincipalname Jane@domain.com -FilterRelevantForTeams
```

Displays all relevant Teams licenses assigned to Jane@domain.com

### EXAMPLE 4
```
Import-Csv User.csv | Get-AzureAdUserLicense
```

Displays all licenses assigned to Users from User.csv, Column UserPrincipalname, ObjectId or Identity.
The input file must have a single column heading of "UserPrincipalname" with properly formatted UPNs.

## PARAMETERS

### -UserPrincipalName
The UserPrincipalname or ObjectId of the Object.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: ObjectId, Identity

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -FilterRelevantForTeams
Filters the output and displays only Licenses relevant to Teams

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String
## OUTPUTS

### System.Object
## NOTES
Requires a connection to Azure Active Directory

## RELATED LINKS

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-AzureAdUserLicense.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Get-AzureAdUserLicense.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_Licensing.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_Licensing.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_UserManagement.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_UserManagement.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/)

