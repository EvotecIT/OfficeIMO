---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Optimize-OfficePdf
## SYNOPSIS
Writes a smaller or sanitized copy of a PDF.

## SYNTAX
### Path (Default)
```powershell
Optimize-OfficePdf [-Path] <string[]> -OutputPath <string> [-NoCompressStream] [-KeepUnreferencedObject] [-KeepDuplicateStream] [-AllowLarger] [-MinimumStreamCompressionSize <Int32>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Optimize-OfficePdf [-Path] <string[]> -OutputDirectory <string> [-NoCompressStream] [-KeepUnreferencedObject] [-KeepDuplicateStream] [-AllowLarger] [-MinimumStreamCompressionSize <Int32>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathSanitize
```powershell
Optimize-OfficePdf [-Path] <string[]> -OutputPath <string> -Sanitize [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-HiddenDataOptions <OfficeHiddenDataOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathSanitizeDirectory
```powershell
Optimize-OfficePdf [-Path] <string[]> -OutputDirectory <string> -Sanitize [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-HiddenDataOptions <OfficeHiddenDataOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Optimize-OfficePdf -OutputPath <string> -LiteralPath <string[]> [-NoCompressStream] [-KeepUnreferencedObject] [-KeepDuplicateStream] [-AllowLarger] [-MinimumStreamCompressionSize <Int32>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathSanitize
```powershell
Optimize-OfficePdf -OutputPath <string> -Sanitize -LiteralPath <string[]> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-HiddenDataOptions <OfficeHiddenDataOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Optimize-OfficePdf -OutputDirectory <string> -LiteralPath <string[]> [-NoCompressStream] [-KeepUnreferencedObject] [-KeepDuplicateStream] [-AllowLarger] [-MinimumStreamCompressionSize <Int32>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathSanitizeDirectory
```powershell
Optimize-OfficePdf -OutputDirectory <string> -Sanitize -LiteralPath <string[]> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-HiddenDataOptions <OfficeHiddenDataOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Writes a smaller or sanitized copy of a PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Optimize-OfficePdf -Path .\Reports\Annual-Report.pdf -OutputPath .\Reports\Annual-Report-Small.pdf
```


## PARAMETERS

### -AllowLarger
AllowLarger for this operation.

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

### -IgnorePermissionRestrictions
After successful authentication only; never bypasses a missing password.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory, LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepDuplicateStream
KeepDuplicateStream for this operation.

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

### -KeepUnreferencedObject
KeepUnreferencedObject for this operation.

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

### -LiteralPath
LiteralPath for this operation.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
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
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory, LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinimumStreamCompressionSize
MinimumStreamCompressionSize for this operation.

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

### -NoClobber
NoClobber for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory, LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCompressStream
NoCompressStream for this operation.

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

### -OutputDirectory
OutputDirectory for this operation.

```yaml
Type: String
Parameter Sets: PathDirectory, PathSanitizeDirectory, LiteralPathDirectory, LiteralPathSanitizeDirectory
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
Parameter Sets: Path, PathSanitize, LiteralPath, LiteralPathSanitize
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
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory, LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
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
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory, LiteralPath, LiteralPathSanitize, LiteralPathDirectory, LiteralPathSanitizeDirectory
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
Parameter Sets: Path, PathDirectory, PathSanitize, PathSanitizeDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HiddenDataOptions
Hidden-data categories that -Sanitize removes, from New-OfficeHiddenDataOptions or a hashtable such as @{ Kinds = 'ActiveContent, Attachments, DocumentProperties' }. Defaults to attachments and active content: scripts, unsafe actions and links, and rich media.

```yaml
Type: OfficeHiddenDataOptions
Parameter Sets: PathSanitize, PathSanitizeDirectory, LiteralPathSanitize, LiteralPathSanitizeDirectory
Aliases: SanitizationOptions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sanitize
Sanitize for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: PathSanitize, PathSanitizeDirectory, LiteralPathSanitize, LiteralPathSanitizeDirectory
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
