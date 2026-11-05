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
### Context (Default)
```powershell
Set-OfficePdfMetadata [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfMetadata -Document <PdfDocument> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficePdfMetadata [-Path] <string[]> -OutputPath <string> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Set-OfficePdfMetadata [-Path] <string[]> -OutputDirectory <string> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Set-OfficePdfMetadata -LiteralPath <string[]> -OutputPath <string> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Set-OfficePdfMetadata -LiteralPath <string[]> -OutputDirectory <string> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Incremental] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
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

### -Author
Document author; an empty string clears it.

```yaml
Type: String
Parameter Sets: Context, Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Explicit caller-owned document for inspection or editing.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After authentication, explicitly ignore owner-imposed usage restrictions.

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

### -Incremental
Append a metadata revision instead of rewriting existing PDF content.

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

### -Keywords
Document keywords; an empty string clears them.

```yaml
Type: String
Parameter Sets: Context, Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal input files, including FileInfo.FullName.

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
Native source parsing, authentication and resource settings.

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

### -OutputDirectory
Destination folder retaining input file names.

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

### -PassThru
Emit the updated native document or each saved FileInfo.

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
Source password.

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
Input files with FileSystem wildcard expansion.

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

### -Subject
Document subject; an empty string clears it.

```yaml
Type: String
Parameter Sets: Context, Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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
Parameter Sets: Context, Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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
