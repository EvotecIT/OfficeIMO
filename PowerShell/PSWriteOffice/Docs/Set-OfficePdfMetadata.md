---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfMetadata
## SYNOPSIS
Sets PDF document metadata on generated documents or existing PDF files.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfMetadata [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfMetadata -Document <PdfDocument> [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### File
```powershell
Set-OfficePdfMetadata -Path <string> -OutputPath <string> [-Password <string>] [-IgnorePermissionRestrictions] [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-PassThru] [-Incremental] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
In a New-OfficePdf script block this command updates the generated document metadata.
With -Path and -OutputPath, it rewrites an existing PDF with updated metadata unless -Incremental is used.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfMetadata -Title 'Service Review' -Author 'PSWriteOffice' -Subject 'Operations'
  PdfHeading 'Service Review'
}
```

Stores metadata on a newly generated PDF.

### EXAMPLE 2
```powershell
PS> Set-OfficePdfMetadata -Path .\Input.pdf -OutputPath .\Output.pdf -Title 'Reviewed package' -Author 'Operations'
```

Writes a new PDF with updated metadata.

## PARAMETERS

### -Author
Document author.

```yaml
Type: String
Parameter Sets: Context, Document, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
PDF document to update outside the DSL context.

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
After successful password authentication, explicitly ignore owner-imposed metadata-modification restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Incremental
Append a metadata-only incremental PDF revision instead of rewriting the existing PDF bytes.

```yaml
Type: SwitchParameter
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Keywords
Document keywords.

```yaml
Type: String
Parameter Sets: Context, Document, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Output PDF path when rewriting an existing PDF.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to authenticate an encrypted PDF.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Existing PDF path to rewrite with updated metadata.

```yaml
Type: String
Parameter Sets: File
Aliases: FilePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Document subject.

```yaml
Type: String
Parameter Sets: Context, Document, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Document title.

```yaml
Type: String
Parameter Sets: Context, Document, File
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

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
