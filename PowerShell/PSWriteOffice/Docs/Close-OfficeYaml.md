---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeYaml
## SYNOPSIS
Completes a live YAML document's lifetime without saving.

## SYNTAX
### Document (Default)
```powershell
Close-OfficeYaml -Document <YamlDocument> [<CommonParameters>]
```

## DESCRIPTION
Completes a live YAML document's lifetime without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeYaml -Document $document
```


### EXAMPLE 2
```powershell
$document = Open-OfficeYaml -Path ./values.yaml
(Get-OfficeYamlNode -Document $document -Select 'image.tag').Value
Close-OfficeYaml -Document $document
```


## PARAMETERS

### -Document
YAML document to close.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Yaml.YamlDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
