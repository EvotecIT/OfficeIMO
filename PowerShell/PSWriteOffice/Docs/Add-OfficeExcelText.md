---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelText
## SYNOPSIS
Adds a styled rich-text run to the enclosing cell or comment.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelText [-Text] <string> [-Bold] [-Italic] [-Strikethrough] [-Underline <ExcelUnderlineStyle>] [-Color <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-VerticalAlignment <ExcelVerticalTextAlignment>] [<CommonParameters>]
```

## DESCRIPTION
Adds a styled rich-text run to the enclosing cell or comment.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Status.xlsx {
    ExcelSheet 'Summary' {
        ExcelCell A1 {
            ExcelText 'Owner: ' -Bold
            ExcelText 'Infrastructure team'
        }
    }
}
```

Builds a cell value from a bold label and plain text.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Migration.xlsx {
    ExcelSheet 'Overview' {
        ExcelCell B3 {
            ExcelText 'Cutover risk: '
            ExcelText 'High' -Bold -Color Danger -Size 12pt
            ExcelText '1' -VerticalAlignment Superscript
        }
        ExcelCell B20 '1 Vendor patch not yet certified for production.'
    }
}
```

Colours the risk word by its theme role and adds a superscript reference to the notes below.

## PARAMETERS

### -Bold
Applies bold text. An explicit false value clears bold; omission preserves the native value.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Text or rule colour expressed as a literal OfficeColor or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Applies italic text. An explicit false value clears italic; omission preserves the native value.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Font size, interpreted as points for numbers; explicit physical units are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Applies strikethrough; explicit false clears it.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Comment text.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Apply or clear underline styling across the target row span.

```yaml
Type: ExcelUnderlineStyle
Parameter Sets: Default
Aliases: None
Possible values: Single, Double, SingleAccounting, DoubleAccounting, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalAlignment
Normal, superscript or subscript positioning for the text run.

```yaml
Type: ExcelVerticalTextAlignment
Parameter Sets: Default
Aliases: None
Possible values: Baseline, Superscript, Subscript

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

- `None`

## RELATED LINKS

- None
