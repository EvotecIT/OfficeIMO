---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficePdfLayoutOverlay
## SYNOPSIS
Exports native PDF word, line, region and reading-order diagnostics as PNG or SVG.

## SYNTAX
### Path (Default)
```powershell
Export-OfficePdfLayoutOverlay [-Path] <string[]> -OutputPath <string> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <Object[]>] [-LayoutOptions <PdfTextLayoutOptions>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Export-OfficePdfLayoutOverlay [-Path] <string[]> -OutputDirectory <string> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <Object[]>] [-LayoutOptions <PdfTextLayoutOptions>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficePdfLayoutOverlay -OutputPath <string> -LiteralPath <string[]> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <Object[]>] [-LayoutOptions <PdfTextLayoutOptions>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Export-OfficePdfLayoutOverlay -OutputDirectory <string> -LiteralPath <string[]> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <Object[]>] [-LayoutOptions <PdfTextLayoutOptions>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports native PDF word, line, region and reading-order diagnostics as PNG or SVG.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficePdfLayoutOverlay -LiteralPath './Report[1].pdf' -OutputPath ./Report-layout.svg -Page 1 -Format Svg -PassThru
```


## PARAMETERS

### -Format
Format setting.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
IgnorePermissionRestrictions setting.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutOptions
LayoutOptions setting.

```yaml
Type: PdfTextLayoutOptions
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal input PDF paths; accepts FileInfo.FullName.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathDirectory
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native parsing, authentication and resource policy.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
NoClobber setting.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Options setting.

```yaml
Type: Object[]
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Batch destination retaining source names.

```yaml
Type: String
Parameter Sets: PathDirectory, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single-file destination.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Page setting.

```yaml
Type: Int32
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return native annotation edit reports.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password setting.

```yaml
Type: SecureString
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input PDF paths with wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path, PathDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Scale
Scale setting.

```yaml
Type: Double
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
