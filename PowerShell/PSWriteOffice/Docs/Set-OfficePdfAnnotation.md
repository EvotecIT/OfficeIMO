---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfAnnotation
## SYNOPSIS
Updates native annotation contents, styling and actions by identity or target URI.

## SYNTAX
### Document (Default)
```powershell
Set-OfficePdfAnnotation -ObjectNumber <int> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficePdfAnnotation [-Path] <string[]> -OutputPath <string> -ObjectNumber <int> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Set-OfficePdfAnnotation [-Path] <string[]> -ObjectNumber <int> -OutputDirectory <string> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathUri
```powershell
Set-OfficePdfAnnotation [-Path] <string[]> -OutputPath <string> -Uri <string> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathUriDirectory
```powershell
Set-OfficePdfAnnotation [-Path] <string[]> -OutputDirectory <string> -Uri <string> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Set-OfficePdfAnnotation -OutputPath <string> -ObjectNumber <int> -LiteralPath <string[]> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathUri
```powershell
Set-OfficePdfAnnotation -OutputPath <string> -Uri <string> -LiteralPath <string[]> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Set-OfficePdfAnnotation -ObjectNumber <int> -OutputDirectory <string> -LiteralPath <string[]> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathUriDirectory
```powershell
Set-OfficePdfAnnotation -OutputDirectory <string> -Uri <string> -LiteralPath <string[]> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentUri
```powershell
Set-OfficePdfAnnotation -Uri <string> [-Options <PdfAnnotationUpdateOptions>] [-Contents <string>] [-Title <string>] [-Name <string>] [-Flags <Int32>] [-Color <OfficeColor>] [-RemoveAction] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Native confirmation runs before loading or editing a file. Live inputs stay caller-owned; a replacement document is adopted by the enclosing content block or returned for explicit assignment. PassThru returns native edit reports.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfAnnotation -ObjectNumber 1
```


### EXAMPLE 2
```powershell
Set-OfficePdfAnnotation -Uri 'https://example.com'
```


### EXAMPLE 3
```powershell
Set-OfficePdfAnnotation -OutputPath 'C:\Path' -ObjectNumber 1 -LiteralPath @('C:\Path')
```


## PARAMETERS

### -Color
Replacement annotation RGB color.

```yaml
Type: OfficeColor
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Contents
Contents setting.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Native caller-owned PDF; omission uses the enclosing content block.

```yaml
Type: PdfDocument
Parameter Sets: Document, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Flags
Raw PDF annotation flags.

```yaml
Type: Int32
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

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
Parameter Sets: Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory
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
Parameter Sets: LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory
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
Parameter Sets: Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Name setting.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ObjectNumber
ObjectNumber setting.

```yaml
Type: Int32
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Native annotation edit policy and fields. Explicit parameters override this captured policy; append-only residual data remains opt-in.

```yaml
Type: PdfAnnotationUpdateOptions
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
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
Parameter Sets: PathDirectory, PathUriDirectory, LiteralPathDirectory, LiteralPathUriDirectory
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
Parameter Sets: Path, PathUri, LiteralPath, LiteralPathUri
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return native annotation edit reports.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
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
Parameter Sets: Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory
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
Parameter Sets: Path, PathDirectory, PathUri, PathUriDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RemoveAction
RemoveAction setting.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Title setting.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, PathUri, PathUriDirectory, LiteralPath, LiteralPathUri, LiteralPathDirectory, LiteralPathUriDirectory, DocumentUri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uri
Exact or wildcard target URI used to select link annotations.

```yaml
Type: String
Parameter Sets: PathUri, PathUriDirectory, LiteralPathUri, LiteralPathUriDirectory, DocumentUri
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

- `System.Int32`
- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfDocument`
- `OfficeIMO.Pdf.PdfAnnotationEditResult`

## RELATED LINKS

- None
