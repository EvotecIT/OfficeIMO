---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficePdf
## SYNOPSIS
Flattens forms or annotations, or adds a searchable OCR text layer, writing a new PDF.

## SYNTAX
### Flatten (Default)
```powershell
Update-OfficePdf [-Path] <string[]> -OutputPath <string> [-FlattenForm] [-FlattenAnnotation] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### FlattenDirectory
```powershell
Update-OfficePdf [-Path] <string[]> -OutputDirectory <string> [-FlattenForm] [-FlattenAnnotation] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Searchable
```powershell
Update-OfficePdf [-Path] <string[]> -OutputPath <string> -Searchable [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-OcrOptions <OfficeOcrOptions>] [-OcrMergeOptions <PdfOcrMergeOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SearchableDirectory
```powershell
Update-OfficePdf [-Path] <string[]> -OutputDirectory <string> -Searchable [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-OcrOptions <OfficeOcrOptions>] [-OcrMergeOptions <PdfOcrMergeOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### FlattenLiteralPath
```powershell
Update-OfficePdf -OutputPath <string> -LiteralPath <string[]> [-FlattenForm] [-FlattenAnnotation] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SearchableLiteralPath
```powershell
Update-OfficePdf -OutputPath <string> -LiteralPath <string[]> -Searchable [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-OcrOptions <OfficeOcrOptions>] [-OcrMergeOptions <PdfOcrMergeOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### FlattenLiteralPathDirectory
```powershell
Update-OfficePdf -OutputDirectory <string> -LiteralPath <string[]> [-FlattenForm] [-FlattenAnnotation] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SearchableLiteralPathDirectory
```powershell
Update-OfficePdf -OutputDirectory <string> -LiteralPath <string[]> -Searchable [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-OcrOptions <OfficeOcrOptions>] [-OcrMergeOptions <PdfOcrMergeOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Flattens forms or annotations, or adds a searchable OCR text layer, writing a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Update-OfficePdf -Path .\Forms\Access-Request-Ada.pdf -OutputPath .\Forms\Access-Request-Ada-Final.pdf -FlattenForm
```


## PARAMETERS

### -AppearanceFont
AppearanceFont for this operation.

```yaml
Type: String
Parameter Sets: Flatten, FlattenDirectory, FlattenLiteralPath, FlattenLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppearanceFontPath
AppearanceFontPath for this operation.

```yaml
Type: String
Parameter Sets: Flatten, FlattenDirectory, FlattenLiteralPath, FlattenLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FlattenAnnotation
FlattenAnnotation for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Flatten, FlattenDirectory, FlattenLiteralPath, FlattenLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FlattenForm
FlattenForm for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Flatten, FlattenDirectory, FlattenLiteralPath, FlattenLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful authentication only; never bypasses a missing password.

```yaml
Type: SwitchParameter
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory, FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this operation.

```yaml
Type: String[]
Parameter Sets: FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
LoadOptions for this operation.

```yaml
Type: PdfLoadOptions
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory, FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
NoClobber for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory, FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OcrMergeOptions
Replaces -PdfOptions.

```yaml
Type: PdfOcrMergeOptions
Parameter Sets: Searchable, SearchableDirectory, SearchableLiteralPath, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OcrOptions
New-OfficeOcrOptions: engine, languages, Tesseract path, data directory, DPI, confidence; same type Read-/Convert-OfficeDocument -Ocr use.

```yaml
Type: OfficeOcrOptions
Parameter Sets: Searchable, SearchableDirectory, SearchableLiteralPath, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
OutputDirectory for this operation.

```yaml
Type: String
Parameter Sets: FlattenDirectory, SearchableDirectory, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
OutputPath for this operation.

```yaml
Type: String
Parameter Sets: Flatten, Searchable, FlattenLiteralPath, SearchableLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
PassThru for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory, FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).

```yaml
Type: SecureString
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory, FlattenLiteralPath, SearchableLiteralPath, FlattenLiteralPathDirectory, SearchableLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for this operation.

```yaml
Type: String[]
Parameter Sets: Flatten, FlattenDirectory, Searchable, SearchableDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Searchable
Searchable for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Searchable, SearchableDirectory, SearchableLiteralPath, SearchableLiteralPathDirectory
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

- `System.String[]`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
