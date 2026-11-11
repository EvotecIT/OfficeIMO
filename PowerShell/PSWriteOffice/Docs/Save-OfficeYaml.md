---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeYaml
## SYNOPSIS
Saves a live YAML document and keeps it available.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeYaml [[-Path] <string>] -Document <YamlDocument> [-WriteOptions <YamlSaveOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The default save is minimal-change: untouched text is copied from the source, in the encoding and byte-order
mark it was loaded with, and only edited nodes are written anew. WriteOptions with Reformat rewrites the whole
document through the writer.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeYaml -Document $document
```


### EXAMPLE 2
```powershell
Save-OfficeYaml -Document $document -Path ./out/values.yaml -WriteOptions @{ Reformat = $true; WriterOptions = [OfficeIMO.Yaml.YamlWriterOptions]@{ MappingIndent = 4 } } -PassThru
```


## PARAMETERS

### -Document
YAML document to save.

```yaml
Type: YamlDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the saved file.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional destination, which becomes the document's FilePath; omission saves to its FilePath.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Native save settings or a hashtable of their properties.

```yaml
Type: YamlSaveOptions
Parameter Sets: Document
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

- `OfficeIMO.Yaml.YamlDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
