---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficeToml
## SYNOPSIS
Applies a JSON Patch (RFC 6902) or JSON Merge Patch (RFC 7386) to a live TOML document.

## SYNTAX
### JsonPatch (Default)
```powershell
Update-OfficeToml [-JsonPatch] <Object> [-PassThru] [-Document <TomlDocument>] [<CommonParameters>]
```

### MergePatch
```powershell
Update-OfficeToml -MergePatch <Object> [-PassThru] [-Document <TomlDocument>] [<CommonParameters>]
```

## DESCRIPTION
-JsonPatch takes a list of add, remove, replace, move, copy and test operations whose paths are JSON Pointers
such as /project/version. -MergePatch takes a partial document whose keys overwrite the document's, where $null
removes a key. Either can be JSON text or PowerShell hashtables, arrays and objects; dates in PowerShell values are
written as text.

Only the patched lines change; comments and layout elsewhere stay. New objects become [header] tables and
arrays of objects [[header]] arrays of tables. TOML has no null, so a null value in a JSON Patch fails. A patch is
all or nothing: when a test fails, a path is missing or an edit cannot be made, a non-terminating toml-patch-failed
error is written and the document is unchanged.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeToml -Path ./pyproject.toml {
    Update-OfficeToml -JsonPatch @(
        @{ op = 'test'; path = '/project/version'; value = '1.4.0' }
        @{ op = 'replace'; path = '/project/version'; value = '1.5.0' }
    )
}
```


### EXAMPLE 2
```powershell
Update-OfficeToml -Document $document -MergePatch @{ tool = @{ ruff = @{ 'line-length' = 100; exclude = $null } } } -PassThru
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: TomlDocument
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
The JSON Patch operations: JSON text, or an array of hashtables or objects.

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
The merge patch: JSON text, or a hashtable or object; $null values remove keys.

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

- `OfficeIMO.Toml.TomlDocument`

## OUTPUTS

- `OfficeIMO.Toml.TomlPatchReport`

## RELATED LINKS

- None
