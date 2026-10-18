---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointLayoutBox
## SYNOPSIS
Computes reusable layout boxes for a presentation.

## SYNTAX
### Content (Default)
```powershell
Get-OfficePowerPointLayoutBox [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [<CommonParameters>]
```

### Columns
```powershell
Get-OfficePowerPointLayoutBox -ColumnCount <int> [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [-GutterCm <double>] [<CommonParameters>]
```

### Rows
```powershell
Get-OfficePowerPointLayoutBox -RowCount <int> [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [-GutterCm <double>] [<CommonParameters>]
```

## DESCRIPTION
Returns the content box for a slide or equal column/row boxes derived from the current slide size.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointLayoutBox.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    $box = Get-OfficePowerPointLayoutBox -MarginCm 1.5
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Inside the content box' -X ($box.LeftPoints) -Y ($box.TopPoints) -Width ($box.WidthPoints) -Height 60
}
```

Returns a content box and uses it to position slide text.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointColumns.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    $columns = Get-OfficePowerPointLayoutBox -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Left column' -X ($columns[0].LeftPoints) -Y ($columns[0].TopPoints) -Width ($columns[0].WidthPoints) -Height 80
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Right column' -X ($columns[1].LeftPoints) -Y ($columns[1].TopPoints) -Width ($columns[1].WidthPoints) -Height 80
}
```

Uses two layout boxes to place text in columns.

## PARAMETERS

### -ColumnCount
Number of columns to generate.

```yaml
Type: Int32
Parameter Sets: Columns
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GutterCm
Column or row gutter in centimeters.

```yaml
Type: Double
Parameter Sets: Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginCm
Outer slide margin in centimeters.

```yaml
Type: Double
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to inspect (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -RowCount
Number of rows to generate.

```yaml
Type: Int32
Parameter Sets: Rows
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointLayoutBox`

## RELATED LINKS

- None
