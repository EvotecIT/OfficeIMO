---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeYamlNode
## SYNOPSIS
Removes mapping entries or sequence items from a live YAML document by path.

## SYNTAX
### Document (Default)
```powershell
Remove-OfficeYamlNode [-Select] <string[]> [-Document <YamlDocument>] [<CommonParameters>]
```

## DESCRIPTION
An entry goes with its own comment lines and line; neighboring entries and section comments stay. A path with
nothing at it writes a non-terminating ObjectNotFound error.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeYaml -Path ./values.yaml {
    Remove-OfficeYamlNode -Select 'ingress.annotations["nginx.ingress.kubernetes.io/rewrite-target"]'
}
```


### EXAMPLE 2
```powershell
Remove-OfficeYamlNode -Document $document -Select 'spec.ports[1]' -ErrorAction SilentlyContinue
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: YamlDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Select
Paths of the entries or items to remove.

```yaml
Type: String[]
Parameter Sets: Document
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

- `OfficeIMO.Yaml.YamlDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
