---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeExcelRangeImage
## SYNOPSIS
Exports one worksheet range as an image file.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeExcelRangeImage [-Path] <string> [[-OutputPath] <string>] -WorksheetName <string> -Range <string> [-Format <OfficeImageExportFormat>] [-Options <ExcelImageExportOptions>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeExcelRangeImage [[-OutputPath] <string>] -Document <ExcelDocument> -WorksheetName <string> -Range <string> [-Format <OfficeImageExportFormat>] [-Options <ExcelImageExportOptions>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports one worksheet range as an image file.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeExcelRangeImage -Path .\Report.xlsx -WorksheetName Summary -Range A1:H20 -OutputPath .\summary.png
```


## PARAMETERS

### -Document
Open workbook instance.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Force
Replace an existing destination file.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, Document
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional rendering, size, font, and diagnostic policy settings.

```yaml
Type: ExcelImageExportOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional destination image file. When omitted, returns the in-memory image result only.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the structured image export result when a destination path is used.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 range address to export.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName
Name of the worksheet containing the range.

```yaml
Type: String
Parameter Sets: Path, Document
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
