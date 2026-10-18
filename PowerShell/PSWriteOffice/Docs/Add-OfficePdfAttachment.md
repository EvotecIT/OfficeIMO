---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfAttachment
## SYNOPSIS
Adds an embedded file attachment to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfAttachment [-Path] <string> [-Name <string>] [-MimeType <string>] [-Relationship <PdfAssociatedFileRelationship>] [-Description <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfAttachment [-Path] <string> -Document <PdfDocument> [-Name <string>] [-MimeType <string>] [-Relationship <PdfAssociatedFileRelationship>] [-Description <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an embedded file attachment to a generated PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $dataPath = '.\Examples\Documents\ServiceData.json'
Set-Content -Path $dataPath -Value '{ "service": "Directory", "status": "Healthy" }'
New-OfficePdf -Path .\Examples\Documents\PdfWithAttachment.pdf {
    Add-OfficePdfHeading -Text 'Service report'
    Add-OfficePdfAttachment -Path $dataPath -Name 'service-data.json' -MimeType 'application/json' -Description 'Source data used by the report.'
}
```

Embeds a supporting JSON file in the generated PDF.

## PARAMETERS

### -Description
Optional human-readable attachment description.

```yaml
Type: String
Parameter Sets: Context, Document
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

### -MimeType
Optional MIME type for the embedded file.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional embedded file name. The source file name is used when omitted.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Accepted for compatibility. The replacement document is always emitted when -Document is used.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
File path to embed in the generated PDF.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Relationship
Associated-file relationship between the PDF and the embedded file.

```yaml
Type: PdfAssociatedFileRelationship
Parameter Sets: Context, Document
Aliases: None
Possible values: Unspecified, Source, Data, Alternative, Supplement, C2paManifest

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

## RELATED LINKS

- None
