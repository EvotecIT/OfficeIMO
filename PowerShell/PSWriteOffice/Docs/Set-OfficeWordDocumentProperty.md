---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordDocumentProperty
## SYNOPSIS
Sets a built-in or custom document property on a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordDocumentProperty [-Name] <string> [[-Value] <Object>] [-Document <WordDocument>] [-Custom] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a built-in or custom document property on a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\QuarterlyReport.docx {
    Set-OfficeWordDocumentProperty -Name Title -Value 'Quarterly Report'
    Set-OfficeWordDocumentProperty -Name ReleaseStatus -Value 'Approved' -Custom
    Add-OfficeWordParagraph -Text 'Approved quarterly report'
}
Get-OfficeWordDocumentProperty -Path .\QuarterlyReport.docx -Name Title, ReleaseStatus
```

Writes document metadata during composition and reads it back for proof.

## PARAMETERS

### -Custom
Treat the property as a custom document property.

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

### -Name
Property name to update.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
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

### -Value
Property value.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
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
