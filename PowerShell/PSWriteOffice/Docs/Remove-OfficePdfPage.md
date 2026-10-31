---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficePdfPage
## SYNOPSIS
Removes selected pages from a live PDF or writes edited file copies.

## SYNTAX
### Document (Default)
```powershell
Remove-OfficePdfPage -Page <PdfPageSelector> [-Document <PdfDocument>] [-AllowLoss] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Remove-OfficePdfPage [-Path] <string[]> -Page <PdfPageSelector> [-OutputPath <string>] [-OutputDirectory <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-AllowLoss] [-PassThru] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Remove-OfficePdfPage -Page <PdfPageSelector> -LiteralPath <string[]> [-OutputPath <string>] [-OutputDirectory <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-AllowLoss] [-PassThru] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Inside Open-OfficePdf, omission of Document updates the current document. Outside a block, PassThru returns the native edit for assignment or further piping. File outputs are published atomically.

## EXAMPLES

### EXAMPLE 1
```powershell
Remove-OfficePdfPage -Path ./Report.pdf -Page '4-5' -OutputPath ./Report-Clean.pdf -PassThru
```


### EXAMPLE 2
```powershell
Open-OfficePdf ./Report.pdf { Remove-OfficePdfPage -Page last }
```


## PARAMETERS

### -AllowLoss
Explicitly accept fidelity loss when saving file outputs.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Explicit native input; omission uses the enclosing New-/Open- content block.

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
After authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native source parsing, authentication and resource-limit settings.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Refuse replacement of an existing file output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder for one or more inputs, retaining their file names.

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

### -OutputPath
Single-file destination. Cannot be combined with OutputDirectory.

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

### -Page
Native selector, one-based page or range expression, including last, odd and even.

```yaml
Type: PdfPageSelector
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated native document or each saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Source password as secure input or a warned plaintext value.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets input paths with provider wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ThrottleLimit
Gets or sets the maximum number of concurrent background operations.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

- `OfficeIMO.Pdf.PdfDocument`
- `System.String[]`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
