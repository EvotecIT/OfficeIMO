---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeOneNote
## SYNOPSIS
Reads an offline OneNote section, notebook hierarchy, or packaged notebook.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeOneNote [-Path] <string> [-Options <Object[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeOneNote -LiteralPath <string> [-Options <Object[]>] [<CommonParameters>]
```

## DESCRIPTION
Reads an offline OneNote section, notebook hierarchy, or packaged notebook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $section = Open-OfficeOneNote -Path .\Operations.one
$section.Pages | Select-Object Title, CreatedUtc, LastModifiedUtc
```


## PARAMETERS

### -LiteralPath
Exact input path without wildcard expansion.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Bounded section and revision-store read options.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path
Aliases: None
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
