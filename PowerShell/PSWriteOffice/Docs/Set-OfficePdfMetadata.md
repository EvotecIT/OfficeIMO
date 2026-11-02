---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfMetadata
## SYNOPSIS
Updates metadata on a generated PDF, an open editing block, or file copies.

## SYNTAX
### Document (Default)
```powershell
Set-OfficePdfMetadata [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-Document <PdfDocument>] [-AllowLoss] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficePdfMetadata [-Path] <string[]> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-OutputPath <string>] [-OutputDirectory <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-AllowLoss] [-PassThru] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Set-OfficePdfMetadata -LiteralPath <string[]> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-OutputPath <string>] [-OutputDirectory <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-AllowLoss] [-PassThru] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Omitted values retain existing metadata; an empty string clears a value. Incremental appends a native metadata revision. File outputs are atomic; document edits return the updated native document with PassThru.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path ./Report.pdf { PdfMetadata -Title 'Service review'; PdfHeading 'Service review' }
```


### EXAMPLE 2
```powershell
Open-OfficePdf ./Report.pdf -OutputPath ./Reviewed.pdf { PdfMetadata -Title 'Reviewed report' }
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

### -Author
Document author; an empty string clears it.

```yaml
Type: String
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

### -Incremental
Append a metadata revision instead of rewriting existing PDF content.

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

### -Keywords
Document keywords; an empty string clears them.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
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

### -Subject
Document subject; an empty string clears it.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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

### -Title
Document title; an empty string clears it.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
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

- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
