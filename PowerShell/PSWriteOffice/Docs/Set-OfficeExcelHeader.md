---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelHeader
## SYNOPSIS
Updates worksheet header text, print flags or an image while retaining the footer.

## SYNTAX
### Text (Default)
```powershell
Set-OfficeExcelHeader [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Left <string>] [-Center <string>] [-Right <string>] [-DifferentFirstPage] [-DifferentOddEven] [-NoAlignWithMargins] [-NoScaleWithDocument] [-PassThru] [<CommonParameters>]
```

### ImagePath
```powershell
Set-OfficeExcelHeader -ImagePath <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Left <string>] [-Center <string>] [-Right <string>] [-DifferentFirstPage] [-DifferentOddEven] [-NoAlignWithMargins] [-NoScaleWithDocument] [-ImagePosition <ExcelHeaderFooterPosition>] [-ImageWidth <OfficeLength>] [-ImageHeight <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### ImageUrl
```powershell
Set-OfficeExcelHeader -ImageUrl <uri> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Left <string>] [-Center <string>] [-Right <string>] [-DifferentFirstPage] [-DifferentOddEven] [-NoAlignWithMargins] [-NoScaleWithDocument] [-ImagePosition <ExcelHeaderFooterPosition>] [-ImageWidth <OfficeLength>] [-ImageHeight <OfficeLength>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates worksheet header text, print flags or an image while retaining the footer.

## EXAMPLES

### EXAMPLE 1
```powershell
ExcelSheet 'Data' { ExcelHeader -Center 'Quarterly review' }
```


## PARAMETERS

### -Center
Center section; Excel field codes such as &P and &N remain active.

```yaml
Type: String
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferentFirstPage
Enable a first-page variant; explicitly false disables it without deleting authored variant text.

```yaml
Type: SwitchParameter
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferentOddEven
Enable odd/even variants; omission retains the setting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ImageHeight
Physical image height; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePath
Literal local image file.

```yaml
Type: String
Parameter Sets: ImagePath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePosition
Image section; defaults to Right.

```yaml
Type: ExcelHeaderFooterPosition
Parameter Sets: ImagePath, ImageUrl
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageUrl
HTTP or HTTPS image URL; existing native remote-image policy applies.

```yaml
Type: Uri
Parameter Sets: ImageUrl
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageWidth
Physical image width; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Left
Left section; empty text clears it, omission preserves it.

```yaml
Type: String
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAlignWithMargins
Disable alignment with page margins; explicitly false restores alignment.

```yaml
Type: SwitchParameter
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoScaleWithDocument
Disable scaling with the printed worksheet; explicitly false restores scaling.

```yaml
Type: SwitchParameter
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated caller-owned worksheet.

```yaml
Type: SwitchParameter
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Right
Right section; empty text clears it, omission preserves it.

```yaml
Type: String
Parameter Sets: Text, ImagePath, ImageUrl
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Text, ImagePath, ImageUrl
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
