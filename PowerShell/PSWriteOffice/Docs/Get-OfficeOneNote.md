---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeOneNote
## SYNOPSIS
Reads an offline OneNote section, notebook hierarchy, or packaged notebook.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeOneNote [-Path] <string> [-Options <OneNoteReaderOptions>] [-NotebookOptions <OneNoteNotebookReaderOptions>] [<CommonParameters>]
```

## DESCRIPTION
Reads an offline OneNote section, notebook hierarchy, or packaged notebook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $section = Get-OfficeOneNote -Path .\Operations.one
$section.Pages | Select-Object Title, CreatedUtc, LastModifiedUtc
```


## PARAMETERS

### -NotebookOptions
Notebook hierarchy, package, and section-error policy.

```yaml
Type: OneNoteNotebookReaderOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Bounded section and revision-store read options.

```yaml
Type: OneNoteReaderOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a .one section, .onetoc2 notebook index, or .onepkg archive.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.OneNote.OneNoteSection`
- `OfficeIMO.OneNote.OneNoteNotebook`

## RELATED LINKS

- None
