---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficeYaml
## SYNOPSIS
Applies a JSON Patch (RFC 6902) or JSON Merge Patch (RFC 7386) to a live YAML document.

## SYNTAX
### JsonPatch (Default)
```powershell
Update-OfficeYaml [-JsonPatch] <Object> [-PassThru] [-Document <YamlDocument>] [<CommonParameters>]
```

### MergePatch
```powershell
Update-OfficeYaml -MergePatch <Object> [-PassThru] [-Document <YamlDocument>] [<CommonParameters>]
```

## DESCRIPTION
-JsonPatch takes a list of add, remove, replace, move, copy and test operations whose paths are JSON Pointers
such as /spec/replicas. -MergePatch takes a partial document whose keys overwrite the document's, where $null removes
a key. Either can be JSON or YAML text, a YAML document, or PowerShell hashtables, arrays and objects.

Only the patched lines change; comments, quoting and layout elsewhere stay. A patch is all or nothing: when a
test fails, a path is missing or an edit cannot be made, a non-terminating yaml-patch-failed error is written and the
document is unchanged. Patches address the first document of a multi-document stream.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeYaml -Path ./deployment.yaml {
    Update-OfficeYaml -JsonPatch @(
        @{ op = 'test'; path = '/spec/replicas'; value = 2 }
        @{ op = 'replace'; path = '/spec/replicas'; value = 3 }
    )
}
```


### EXAMPLE 2
```powershell
Update-OfficeYaml -Document $document -JsonPatch (Get-Content ./scale.patch.json -Raw) -PassThru
```


### EXAMPLE 3
```powershell
Update-OfficeYaml -Document $document -MergePatch @{ image = @{ tag = '1.27.2' }; podAnnotations = $null }
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: YamlDocument
Parameter Sets: JsonPatch, MergePatch
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -JsonPatch
The JSON Patch operations: JSON or YAML text, a YAML document, or an array of hashtables or objects.

```yaml
Type: Object
Parameter Sets: JsonPatch
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MergePatch
The merge patch: JSON or YAML text, a YAML document, or a hashtable or object; $null values remove keys.

```yaml
Type: Object
Parameter Sets: MergePatch
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the report of added, replaced and removed paths.

```yaml
Type: SwitchParameter
Parameter Sets: JsonPatch, MergePatch
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

- `OfficeIMO.Yaml.YamlPatchReport`

## RELATED LINKS

- None
