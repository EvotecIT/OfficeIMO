---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordTable
## SYNOPSIS
Gets tables from a Word document or section.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordTable [-Path] <string> [-IncludeNested] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordTable -Document <WordDocument> [-IncludeNested] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordTable -Section <WordSection> [-IncludeNested] [<CommonParameters>]
```

## DESCRIPTION
Gets tables from a Word document or section.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $tables = Get-OfficeWordTable -Path .\Report.docx -IncludeNested
[pscustomobject]@{
    Path = '.\Report.docx'
    TableCount = @($tables).Count
}
```

Reads tables from the document and includes nested tables when validating a composed report.

## PARAMETERS

### -Document
Document to inspect.

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

### -IncludeNested
Include nested tables.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Section
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the document.

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

### -Section
Section to enumerate.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`

## OUTPUTS

- `OfficeIMO.Word.WordTable`

## RELATED LINKS

- None
