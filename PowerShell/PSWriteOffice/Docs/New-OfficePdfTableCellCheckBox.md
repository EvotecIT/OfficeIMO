---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfTableCellCheckBox
## SYNOPSIS
Creates a typed check box for a PDF table cell.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfTableCellCheckBox [-Name] <string> [-Checked] [-Size <double>] [-CheckedValueName <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a typed check box for a PDF table cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $approved = New-OfficePdfTableCellCheckBox -Name Approved -Checked
$cell = New-OfficePdfTableCell -Text 'Approved' -CheckBox $approved
```

The check box remains an AcroForm field positioned by the OfficeIMO table renderer.

## PARAMETERS

### -Checked
Create the check box in its checked state.

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

### -CheckedValueName
PDF appearance-state name written when checked.

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

### -Name
Unique AcroForm field name.

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

### -Size
Visual square size in PDF points.

```yaml
Type: Double
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

- `OfficeIMO.Pdf.PdfTableCellCheckBox`

## RELATED LINKS

- None
