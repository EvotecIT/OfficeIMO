---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficePdfAnnotation
## SYNOPSIS
Removes native annotations selected by identity, page and subtype.

## SYNTAX
### Document (Default)
```powershell
Remove-OfficePdfAnnotation [-ObjectNumber <Int32>] [-Page <Int32>] [-Subtype <string>] [-KeepPopups] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Remove-OfficePdfAnnotation [-Path] <string[]> -OutputPath <string> [-ObjectNumber <Int32>] [-Page <Int32>] [-Subtype <string>] [-KeepPopups] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Remove-OfficePdfAnnotation [-Path] <string[]> -OutputDirectory <string> [-ObjectNumber <Int32>] [-Page <Int32>] [-Subtype <string>] [-KeepPopups] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Remove-OfficePdfAnnotation -OutputPath <string> -LiteralPath <string[]> [-ObjectNumber <Int32>] [-Page <Int32>] [-Subtype <string>] [-KeepPopups] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Remove-OfficePdfAnnotation -OutputDirectory <string> -LiteralPath <string[]> [-ObjectNumber <Int32>] [-Page <Int32>] [-Subtype <string>] [-KeepPopups] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Native confirmation runs before loading or editing a file. Live inputs stay caller-owned; a replacement document is adopted by the enclosing content block or returned for explicit assignment. PassThru returns native edit reports.

## EXAMPLES

### EXAMPLE 1
```powershell
Remove-OfficePdfAnnotation -Document 'Value'
```


### EXAMPLE 2
```powershell
Remove-OfficePdfAnnotation -OutputPath 'C:\Path' -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Remove-OfficePdfAnnotation -OutputDirectory 'Value' -LiteralPath @('C:\Path')
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

### -KeepPopups
KeepPopups setting.

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

### -ObjectNumber
ObjectNumber setting.

```yaml
Type: Int32
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
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

### -Subtype
Subtype setting.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfDocument`
- `OfficeIMO.Pdf.PdfAnnotationEditResult`

## RELATED LINKS

- None
