---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeAccessTable
## SYNOPSIS
Adds a table to a newly created native database model.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeAccessTable [-Name] <string> -Document <AccessDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a table to a newly created native database model.

## EXAMPLES

### EXAMPLE 1
```powershell
$table = $database | Add-OfficeAccessTable -Name Contacts -PassThru
```


## PARAMETERS

### -Document
Live database document.

```yaml
Type: AccessDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Table name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the created or changed object.

```yaml
Type: SwitchParameter
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

- `OfficeIMO.Access.AccessDocument`

## OUTPUTS

- `OfficeIMO.Access.AccessTable`

## RELATED LINKS

- None
