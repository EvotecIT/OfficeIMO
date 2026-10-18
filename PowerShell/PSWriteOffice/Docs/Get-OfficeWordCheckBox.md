---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordCheckBox
## SYNOPSIS
Gets checkbox content controls from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordCheckBox [-Path] <string> [-Alias <string[]>] [-Tag <string[]>] [-Checked] [-Unchecked] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordCheckBox -Document <WordDocument> [-Alias <string[]>] [-Tag <string[]>] [-Checked] [-Unchecked] [<CommonParameters>]
```

## DESCRIPTION
Gets checkbox content controls from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $openItems = Get-OfficeWordCheckBox -Path .\Report.docx -Tag 'approval-*' -Unchecked
$openItems |
    Select-Object -Property Alias, Tag, IsChecked |
    Format-Table -AutoSize
```

Finds unchecked approval checkboxes by tag so a review workflow can show what is still pending.

## PARAMETERS

### -Alias
Filter by checkbox alias (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Checked
Only return checked checkboxes.

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

### -Tag
Filter by checkbox tag (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Unchecked
Only return unchecked checkboxes.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordCheckBox`

## RELATED LINKS

- None
