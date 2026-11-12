---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeYamlNode
## SYNOPSIS
Sets the value at a path in a live YAML document, creating missing mappings on the way.

## SYNTAX
### Document (Default)
```powershell
Set-OfficeYamlNode [-Select] <string> [-Value] <Object> [-Style <YamlStyle>] [-PassThru] [-Document <YamlDocument>] [<CommonParameters>]
```

## DESCRIPTION
Setting a scalar over a scalar keeps its comments, anchor and tag, and an equal value keeps its original
spelling, so saving changes only the text of what really changed. Hashtables, arrays and objects become mappings
and sequences in the layout of their new neighbors. An index equal to a sequence's length appends.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeYaml -Path ./values.yaml {
    Set-OfficeYamlNode -Select 'image.tag' -Value '1.27.2'
}
```


### EXAMPLE 2
```powershell
Set-OfficeYamlNode -Document $document -Select 'resources.limits' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' })
            Set-OfficeYamlNode -Document $document -Select 'ports[2]' -Value 8443
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

### -PassThru
Return the node that is now at the path.

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

### -Select
Path of the node to set; $ sets the root of the first document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Scalar style for a string value, such as DoubleQuoted or Literal; omission picks a safe style.

```yaml
Type: YamlStyle
Parameter Sets: Document
Aliases: None
Possible values: None, Plain, SingleQuoted, DoubleQuoted, Literal, Folded, Block, Flow

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
The new value: a scalar, hashtable, array, object or YamlNode. $null writes a YAML null.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Yaml.YamlDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
