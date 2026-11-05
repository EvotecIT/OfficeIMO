---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficePdfXfdf
## SYNOPSIS
Exports native PDF form values as XFDF XML or atomically published XFDF files.

## SYNTAX
### Document (Default)
```powershell
Export-OfficePdfXfdf [-NoClobber] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Export-OfficePdfXfdf [-Path] <string[]> [-OutputPath <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Export-OfficePdfXfdf [-Path] <string[]> -OutputDirectory <string> [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficePdfXfdf -LiteralPath <string[]> [-OutputPath <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Export-OfficePdfXfdf -OutputDirectory <string> -LiteralPath <string[]> [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports native PDF form values as XFDF XML or atomically published XFDF files.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficePdfXfdf -Path './Forms/*.pdf' -OutputDirectory ./Values -NoClobber -PassThru
```


## PARAMETERS

### -Document
Native caller-owned PDF; omission uses the enclosing content block.

```yaml
Type: PdfDocument
Parameter Sets: Document
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
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
