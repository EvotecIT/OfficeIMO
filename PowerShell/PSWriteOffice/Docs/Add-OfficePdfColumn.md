---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfColumn
## SYNOPSIS
Defines one column of a PdfRow and its content.

## SYNTAX
### Weight (Default)
```powershell
Add-OfficePdfColumn [-Content] <scriptblock> [-Weight <Double>] [<CommonParameters>]
```

### Width
```powershell
Add-OfficePdfColumn [-Content] <scriptblock> -Width <OfficeLength> [<CommonParameters>]
```

## DESCRIPTION
Defines one column of a PdfRow and its content.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path .\Columns.pdf {
    PdfRow {
        PdfColumn { PdfHeading 'Signals' -Level 2 }
        PdfColumn { PdfPanel 'All checks passed.' }
    }
}
```


## PARAMETERS

### -Content
Content for this operation.

```yaml
Type: ScriptBlock
Parameter Sets: Weight, Width
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Weight
Weight for this operation.

```yaml
Type: Double
Parameter Sets: Weight
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Width
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

- `None`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
