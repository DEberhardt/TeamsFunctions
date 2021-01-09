---
external help file: TeamsFunctions-help.xml
Module Name: TeamsFunctions
online version:
schema: 2.0.0
---

# Test-AzureAdGroup

## SYNOPSIS
Tests whether an Group exists in Azure AD (record found)

## SYNTAX

```
Test-AzureAdGroup [-Identity] <String> [<CommonParameters>]
```

## DESCRIPTION
Simple lookup - does the Group Object exist - to avoid TRY/CATCH statements for processing

## EXAMPLES

### EXAMPLE 1
```
Test-AzureAdGroup -Identity "My Group"
```

Will Return $TRUE only if the object "My Group" is found.
  Will Return $FALSE in any other case

## PARAMETERS

### -Identity
Mandatory.
The Name or User Principal Name (MailNickName) of the Group to test.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.Boolean
## NOTES

## RELATED LINKS

[Find-AzureAdGroup
Find-AzureAdUser
Test-AzureAdGroup
Test-AzureAdUser
Test-TeamsUser]()
