---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfPageNumber
## SYNOPSIS
Inserts a live page number (optionally 'of total') inside PdfHeader or PdfFooter.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfPageNumber [-IncludeTotal] [-Separator <string>] [-Bold] [-Italic] [-Strikethrough] [-Underline <OfficeTextDecorationStyle>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Baseline <PdfTextBaseline>] [<CommonParameters>]
```

## DESCRIPTION
Inserts a live page number (optionally 'of total') inside PdfHeader or PdfFooter.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path .\Numbered.pdf {
    PdfFooter {
        PdfText 'Page '
        PdfPageNumber -IncludeTotal
    }
    PdfParagraph 'Body text.'
}
```


## PARAMETERS

### -BackgroundColor
BackgroundColor for this operation.

```yaml
Type: OfficeColor
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Baseline
Baseline for this operation.

```yaml
Type: PdfTextBaseline
Parameter Sets: Context
Aliases: None
Possible values: Normal, Superscript, Subscript

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Bold for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Color for this operation.

```yaml
Type: OfficeColor
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family name or standard PDF font; completes from the theme and registered fonts.

```yaml
Type: String
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeTotal
IncludeTotal for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Italic for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Separator
Separator for this operation.

```yaml
Type: String
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Size for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Strikethrough for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Shared underline style; bare -Underline means Single.

```yaml
Type: OfficeTextDecorationStyle
Parameter Sets: Context
Aliases: None
Possible values: None, Single, Double, Dotted, Dashed, Wavy, Words

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

- `System.Void`

## RELATED LINKS

- None
