---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficeWordDocument
## SYNOPSIS
Appends one or more Word documents into a base Word document.

## SYNTAX
### Path (Default)
```powershell
Join-OfficeWordDocument [-Path] <string> [-AppendPath] <string[]> [-OutputPath <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Join-OfficeWordDocument [-AppendPath] <string[]> -Document <WordDocument> [-OutputPath <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO.Word document append support and preserves the wrapper as an operator-friendly merge command.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Join-OfficeWordDocument -Path .\Cover.docx -AppendPath .\Body.docx, .\Appendix.docx -OutputPath .\ReleasePacket.docx
    Get-OfficeWordStatistics -Path .\ReleasePacket.docx |
        Select-Object -Property Paragraphs, Tables, Images
)
$proof
```

Appends the source documents with OfficeIMO.Word and then reads back basic structure from the merged output.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Base.docx
$doc | Join-OfficeWordDocument -AppendPath .\Section1.docx, .\Section2.docx -PassThru |
    Save-OfficeWord -Path .\Combined.docx
```

Keeps the wrapper thin by piping the OfficeIMO document object through append and save commands.

## PARAMETERS

### -AppendPath
Documents to append to the base document.

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: SourcePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Base document object.

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

### -Open
Open the saved output with the shell.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path. When omitted for path input, the base document is updated in place.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the merged Word document instead of disposing it.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Base document path.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, BasePath
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

- `OfficeIMO.Word.WordDocument`

## RELATED LINKS

- None
