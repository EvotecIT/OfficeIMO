---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeAccessRow
## SYNOPSIS
Appends PowerShell objects or dictionaries using the declared field types.

## SYNTAX
### Table (Default)
```powershell
Add-OfficeAccessRow [-InputObject] <psobject[]> -Table <AccessTable> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Appends PowerShell objects or dictionaries using the declared field types.

## EXAMPLES

### EXAMPLE 1
```powershell
$rows | Add-OfficeAccessRow -Table $table
```


## PARAMETERS

### -InputObject
One row; omitted properties remain absent, and explicit null remains null.

```yaml
Type: PSObject[]
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the created or changed object.

```yaml
Type: SwitchParameter
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Destination table in a newly created model.

```yaml
Type: AccessTable
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Management.Automation.PSObject[]`

## OUTPUTS

- `OfficeIMO.Access.AccessTable`

## RELATED LINKS

- None
