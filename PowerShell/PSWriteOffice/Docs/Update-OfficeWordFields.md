---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeWordFields
## SYNOPSIS
Updates fields in a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Update-OfficeWordFields [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Refreshes page number fields and queues table-of-contents updates.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx {
    Add-OfficeWordParagraph -Text 'Release report'
    Add-OfficeWordTableOfContents
    Update-OfficeWordFields
}
```

Updates PAGE and NUMPAGES fields and marks TOC fields as dirty before saving.

## PARAMETERS

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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

- `None`

## RELATED LINKS

- None
