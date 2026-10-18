---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordTableOfContents
## SYNOPSIS
Gets the table of contents from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordTableOfContents [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordTableOfContents -Document <WordDocument> [<CommonParameters>]
```

## DESCRIPTION
Gets the table of contents from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$toc = $doc | Get-OfficeWordTableOfContents
if ($toc) {
    $toc | Set-OfficeWordTableOfContents -Text 'Contents' -TextNoContent 'No entries' -PassThru
    $doc | Save-OfficeWord -Path .\Report-Toc.docx
}
```

Gets the TOC from an open document, updates it, and saves a variant.

## PARAMETERS

### -Document
Word document to read.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
Path to the .docx file.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None
