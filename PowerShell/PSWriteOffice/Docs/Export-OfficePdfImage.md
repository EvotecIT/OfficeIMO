---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePdfImage
## SYNOPSIS
Exports PDF pages through the shared PNG, JPEG, TIFF, SVG, or WebP image contract.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficePdfImage [-Path] <string> [-OutputPath] <string> [-PageRange <string>] [-Format <OfficeImageExportFormat>] [-Options <PdfImageExportOptions>] [-ReadOptions <PdfLoadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports PDF pages through the shared PNG, JPEG, TIFF, SVG, or WebP image contract.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\Pages -PageRange '1-3,5'
```

Writes the selected pages. Add -PassThru to receive normalized image results with rendering diagnostics.

## PARAMETERS

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

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

### -Options
Optional DPI, scale, thumbnail, encoding, diagnostics, and resource limits.

```yaml
Type: PdfImageExportOptions
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Optional one-based ranges such as 1-3,5.

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

### -PassThru
Emit one structured image export result per saved page.

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

### -Password
Password used to authenticate an encrypted PDF.

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

### -Path
Path to the PDF.

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

### -ReadOptions
Optional bounded PDF parsing settings.

```yaml
Type: PdfLoadOptions
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

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
