---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeVisioImage
## SYNOPSIS
Exports selected Visio pages through the format-neutral OfficeIMO image pipeline.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeVisioImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-Options <VisioImageExportOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeVisioImage [-OutputPath] <string> -Document <VisioDocument> [-Format <OfficeImageExportFormat>] [-Options <VisioImageExportOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports selected Visio pages through the format-neutral OfficeIMO image pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeVisioImage -Path .\diagram.vsdx -OutputPath .\Images -Format Png
```

Writes one PNG per selected page. Add -PassThru to receive one result object per file.

## PARAMETERS

### -Document
Open Visio document instance.

```yaml
Type: VisioDocument
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
Optional page selection, size, concurrency, and rendering settings.

```yaml
Type: VisioImageExportOptions
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
Destination folder.

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
Emit one structured image export result per saved page.

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
Path to a Visio document.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
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

- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
