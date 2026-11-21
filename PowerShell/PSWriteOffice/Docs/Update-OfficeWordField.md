---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficeWordField
## SYNOPSIS
Updates fields in a Word document.

## SYNTAX
### Default (Default)
```powershell
Update-OfficeWordField [-Document <WordDocument>] [-TableOfContents] [-LayoutOptions <WordPdfExportOptions>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Refreshes page number fields and queues table-of-contents updates.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx {
    Add-OfficeWordParagraph -Text 'Release report'
    Add-OfficeWordTableOfContents
    Update-OfficeWordField
}
```

Updates PAGE and NUMPAGES fields and marks TOC fields as dirty before saving.

### EXAMPLE 2
```powershell
PS> $pdf = New-OfficeWordPdfOptions -Font 'Carlito'
$document = Open-OfficeWord -Path .\Report.docx
$document | Update-OfficeWordField -TableOfContents -LayoutOptions $pdf
Close-OfficeWord -Document $document -Save
```

Regenerates the table of contents and refreshes the supported fields, with page numbers from the same layout and fonts as PDF export with $pdf.

## PARAMETERS

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LayoutOptions
PDF export options whose layout, fonts and resource policy supply page numbers. Implies a full refresh of supported fields.

```yaml
Type: WordPdfExportOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableOfContents
Regenerate tables of contents with page numbers from the PDF layout.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`

## RELATED LINKS

- None
