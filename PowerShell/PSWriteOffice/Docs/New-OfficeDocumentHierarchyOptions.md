---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeDocumentHierarchyOptions
## SYNOPSIS
Creates discoverable token and hierarchy settings for Read-OfficeDocument -As Hierarchy.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeDocumentHierarchyOptions [-MaximumTokens <Int32>] [-OverlapTokens <Int32>] [-MaximumInputChunks <Int32>] [-MaximumOutputChunks <Int32>] [-MaximumHierarchyDepth <Int32>] [-MaximumContextCharacters <Int32>] [-PreferMarkdown] [-IncludeContextInText] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable token and hierarchy settings for Read-OfficeDocument -As Hierarchy.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeDocumentHierarchyOptions -MaximumTokens 500 -OverlapTokens 50 -IncludeContextInText
Read-OfficeDocument -As Hierarchy -Path .\handbook.pdf -HierarchyOptions $options
```


## PARAMETERS

### -IncludeContextInText
Include hierarchy context in chunk text.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumContextCharacters
Maximum heading-context characters retained.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumHierarchyDepth
Maximum heading hierarchy depth.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputChunks
Maximum source chunks accepted.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputChunks
Maximum chunks returned.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTokens
Maximum tokens per output chunk.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverlapTokens
Tokens repeated between adjacent chunks.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreferMarkdown
Prefer Markdown text where the reader supports it.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Reader.ReaderHierarchicalChunkingOptions`

## RELATED LINKS

- None
