---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeWordImage
## SYNOPSIS
Exports one or more Word pages through the format-neutral OfficeIMO image pipeline.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeWordImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-Options <WordImageExportOptions>] [-AllPages] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeWordImage [-OutputPath] <string> -Document <WordDocument> [-Format <OfficeImageExportFormat>] [-Options <WordImageExportOptions>] [-AllPages] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports one or more Word pages through the format-neutral OfficeIMO image pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Report.svg -Format Svg
```

Writes the image quietly. Add -PassThru to receive the structured export result.

### EXAMPLE 2
```powershell
PS> Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Pages -Format Jpeg -AllPages
```

For a bounded batch, create options with New-OfficeWordImageOptions -PageIndex 0 -PageCount 2.

## PARAMETERS

### -AllPages
Export every estimated page to the destination folder.

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

### -Document
Open Word document instance.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
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
Optional page, size, scale, theme, and rendering settings.

```yaml
Type: WordImageExportOptions
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
Destination image file, or destination folder when -AllPages or Options.PageCount requests a batch.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the structured image export result.

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
Path to the Word document.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
