---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfPageBorder
## SYNOPSIS
Sets or clears the generated PDF page border decoration.

## SYNTAX
### Border (Default)
```powershell
Set-OfficePdfPageBorder [-Color <OfficeColor>] [-Width <OfficeLength>] [-Inset <OfficeLength>] [-Opacity <Double>] [-DashStyle <OfficeStrokeDashStyle>] [<CommonParameters>]
```

### Clear
```powershell
Set-OfficePdfPageBorder -Clear [<CommonParameters>]
```

## DESCRIPTION
Sets or clears the generated PDF page border decoration.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPageBorder.pdf {
    Set-OfficePdfPageBorder -Color '#CBD5E1' -Width 0.75 -Inset 24 -Opacity 0.8
    Add-OfficePdfHeading -Text 'Bordered report'
}
```

Decorates generated pages with an OfficeIMO.Pdf page border.

## PARAMETERS

### -Clear
Clear the generated PDF page border decoration.

```yaml
Type: SwitchParameter
Parameter Sets: Clear
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Border color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Border
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DashStyle
Border dash style.

```yaml
Type: OfficeStrokeDashStyle
Parameter Sets: Border
Aliases: None
Possible values: Solid, Dash, Dot, DashDot, DashDotDot

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Inset
Distance from the page edge to the border path in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Border
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Opacity
Border opacity from 0 through 1.

```yaml
Type: Double
Parameter Sets: Border
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Border stroke width in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Border
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

- `None`

## RELATED LINKS

- None
