---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfTableCellField
## SYNOPSIS
Creates a typed text or choice field for a PDF table cell.

## SYNTAX
### Text (Default)
```powershell
New-OfficePdfTableCellField [-Name] <string> [[-Value] <string>] [-Width <double>] [-Height <double>] [-FontSize <double>] [<CommonParameters>]
```

### Choice
```powershell
New-OfficePdfTableCellField [-Name] <string> [[-Value] <string>] -Option <string[]> [-Width <double>] [-Height <double>] [-FontSize <double>] [-ListBox] [<CommonParameters>]
```

## DESCRIPTION
Creates a typed text or choice field for a PDF table cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $reviewer = New-OfficePdfTableCellField -Name Reviewer -Option 'Unassigned', 'Alice', 'Bob' -Value 'Unassigned'
$cell = New-OfficePdfTableCell -Text 'Reviewer' -FormField $reviewer
```

The choice field is positioned by the OfficeIMO PDF table renderer.

## PARAMETERS

### -FontSize
Field font size in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Choice
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Rendered field height in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Choice
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListBox
Render a choice field as a list box instead of a combo box.

```yaml
Type: SwitchParameter
Parameter Sets: Choice
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Unique AcroForm field name.

```yaml
Type: String
Parameter Sets: Text, Choice
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Option
Available values for a choice field.

```yaml
Type: String[]
Parameter Sets: Choice
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Initial field value.

```yaml
Type: String
Parameter Sets: Text, Choice
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rendered field width in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Choice
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

- `OfficeIMO.Pdf.PdfTableCellFormField`

## RELATED LINKS

- None
