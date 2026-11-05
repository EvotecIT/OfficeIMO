---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficePdfXfdf
## SYNOPSIS
Imports bounded DTD-free XFDF into native PDF forms, with atomic file output or a caller-owned live result.

## SYNTAX
### Document (Default)
```powershell
Import-OfficePdfXfdf -Xfdf <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Import-OfficePdfXfdf [-Path] <string[]> -OutputPath <string> -Xfdf <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Import-OfficePdfXfdf [-Path] <string[]> -Xfdf <string> -OutputDirectory <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFile
```powershell
Import-OfficePdfXfdf [-Path] <string[]> -OutputPath <string> -XfdfPath <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFileDirectory
```powershell
Import-OfficePdfXfdf [-Path] <string[]> -OutputDirectory <string> -XfdfPath <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficePdfXfdf -OutputPath <string> -Xfdf <string> -LiteralPath <string[]> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFile
```powershell
Import-OfficePdfXfdf -OutputPath <string> -LiteralPath <string[]> -XfdfPath <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Import-OfficePdfXfdf -Xfdf <string> -OutputDirectory <string> -LiteralPath <string[]> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFileDirectory
```powershell
Import-OfficePdfXfdf -OutputDirectory <string> -LiteralPath <string[]> -XfdfPath <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFile
```powershell
Import-OfficePdfXfdf -XfdfPath <string> [-Options <Object[]>] [-MaxXfdfSize <long>] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Imports bounded DTD-free XFDF into native PDF forms, with atomic file output or a caller-owned live result.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficePdfXfdf -LiteralPath './Form[1].pdf' -XfdfPath ./Values.xfdf -OutputPath ./Completed.pdf
```


## PARAMETERS

### -Document
Native caller-owned PDF; omission uses the enclosing content block.

```yaml
Type: PdfDocument
Parameter Sets: Document, DocumentFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
IgnorePermissionRestrictions setting.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory
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
Parameter Sets: LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory
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
Parameter Sets: Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxXfdfSize
MaxXfdfSize setting.

```yaml
Type: Int64
Parameter Sets: Document, Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory, DocumentFile
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
Parameter Sets: Document, Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory, DocumentFile
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
Parameter Sets: PathDirectory, PathFileDirectory, LiteralPathDirectory, LiteralPathFileDirectory
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the changed, caller-owned PDF document. Dispose it when finished.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory, DocumentFile
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
Parameter Sets: Path, PathDirectory, PathFile, PathFileDirectory, LiteralPath, LiteralPathFile, LiteralPathDirectory, LiteralPathFileDirectory
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
Parameter Sets: Path, PathDirectory, PathFile, PathFileDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Xfdf
Xfdf setting.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -XfdfPath
XfdfPath setting.

```yaml
Type: String
Parameter Sets: PathFile, PathFileDirectory, LiteralPathFile, LiteralPathFileDirectory, DocumentFile
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

- `System.String`
- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
