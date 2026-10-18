---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordReview
## SYNOPSIS
Inspects Word comments, threads, tracked revisions, and unsupported review metadata.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordReview [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordReview -Document <WordDocument> [<CommonParameters>]
```

## DESCRIPTION
Inspects Word comments, threads, tracked revisions, and unsupported review metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeWordReview -Path .\Draft.docx
```

Returns a structured WordReviewReport without changing the document.

## PARAMETERS

### -Document
Open Word document instance.

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
Path to the Word document.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
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

- `OfficeIMO.Word.WordReviewReport`

## RELATED LINKS

- None
