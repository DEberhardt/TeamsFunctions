---
external help file: TeamsFunctions-help.xml
Module Name: TeamsFunctions
online version: https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Test-TeamsResourceAccount.md
schema: 2.0.0
---

# Test-TeamsResourceAccount

## SYNOPSIS
Tests whether an Application Instance exists in Azure AD (record found)

## SYNTAX

```
Test-TeamsResourceAccount [-UserPrincipalName] <String> [-Quick] [<CommonParameters>]
```

## DESCRIPTION
Simple lookup - does the User Object exist - to avoid TRY/CATCH statements for processing

## EXAMPLES

### EXAMPLE 1
```
Test-TeamsResourceAccount -UserPrincipalName "$UPN"
```

Will Return $TRUE only if an CsOnlineApplicationInstance Object with the $UPN is found.
Will Return $FALSE in any other case, including if there is no Connection to AzureAD!

### EXAMPLE 2
```
Test-TeamsResourceAccount -UserPrincipalName "$UPN" -Quick
```

Will Return $TRUE only if an AzureAdObject with the $UPN is found with the Department "Microsoft Communication Application Instance" set)
Will Return $FALSE in any other case, including if there is no Connection to AzureAD!

## PARAMETERS

### -UserPrincipalName
Mandatory.
The sign-in address or User Principal Name of the user account to test.

```yaml
Type: String
Parameter Sets: (All)
Aliases: ObjectId, Identity

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Quick
Optional.
By default, this command queries the CsOnlineApplicationInstance which takes a while.
A cursory check can be performed against the AzureAdUser (Department "Microsoft Communication Application Instance" indicates ResourceAccounts)

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

### Boolean
## NOTES
None

## RELATED LINKS

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Test-TeamsResourceAccount.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/Test-TeamsResourceAccount.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_TeamsResourceAccount.md](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/about_TeamsResourceAccount.md)

[https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/](https://github.com/DEberhardt/TeamsFunctions/tree/master/docs/)

