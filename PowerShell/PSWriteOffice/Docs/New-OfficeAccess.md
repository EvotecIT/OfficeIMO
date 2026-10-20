---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeAccess
## SYNOPSIS
Creates a native Access database model without a provider, seed file, or execution.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeAccess [-Format <AccessFileFormat>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a native Access database model without a provider, seed file, or execution.

## EXAMPLES

### EXAMPLE 1
```powershell
$database = New-OfficeAccess -Format Accdb -Title Contacts
```


## PARAMETERS

### -Format
Native file family.

```yaml
Type: AccessFileFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Mdb, Accdb

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Database title.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Access.AccessDocument`

## RELATED LINKS

- None
