---
external help file: TeamsFunctions-help.xml
Module Name: TeamsFunctions
online version:
schema: 2.0.0
---

# Test-SkypeOnlineConnection

## SYNOPSIS
Tests whether a valid PS Session exists for SkypeOnline (Teams)

## SYNTAX

```
Test-SkypeOnlineConnection [<CommonParameters>]
```

## DESCRIPTION
A connection established via Connect-SkypeOnline is parsed.
This connection must be valid (Available and Opened)

## EXAMPLES

### EXAMPLE 1
```
Test-SkypeOnlineConnection
```

Will Return $TRUE only if a valid and open session is found.

## PARAMETERS

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.Boolean
## NOTES
Added check for Open Session to err on the side of caution.
Use with Disconnect-SkypeOnline when tested negative, then Connect-SkypeOnline

## RELATED LINKS