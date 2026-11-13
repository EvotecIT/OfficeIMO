---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeToml
## SYNOPSIS
Creates a live TOML document, or atomically publishes a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeToml [[-Path] <string>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeToml [-Path] <string> [-Content] <scriptblock> [-Force] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Inside the block, Set-OfficeTomlNode and the other node commands edit the new document without -Document.
Setting a key below a missing table adds the table's [header] section.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeToml -Path ./pyproject.toml {
    Set-OfficeTomlNode -Select 'project.name' -Value 'reporting'
    Set-OfficeTomlNode -Select 'tool.ruff.line-length' -Value 100
}
```


### EXAMPLE 2
```powershell
$document = New-OfficeToml
Set-OfficeTomlNode -Document $document -Select 'title' -Value 'Release notes'
Save-OfficeToml -Document $document -Path ./config.toml
Close-OfficeToml -Document $document
```


## PARAMETERS

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow an existing destination; a failed block preserves its original contents.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: FilePath
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
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

- `None`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Toml.TomlDocument`

## RELATED LINKS

- None
