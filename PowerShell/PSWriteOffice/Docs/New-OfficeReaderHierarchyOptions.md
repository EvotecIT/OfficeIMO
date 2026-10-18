---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeReaderHierarchyOptions
## SYNOPSIS
Creates discoverable token and hierarchy settings for Get-OfficeDocumentHierarchy.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeReaderHierarchyOptions [-MaxTokens <Int32>] [-OverlapTokens <Int32>] [-MaxInputChunks <Int32>] [-MaxOutputChunks <Int32>] [-MaxHierarchyDepth <Int32>] [-MaxContextCharacters <Int32>] [-PreferMarkdown] [-IncludeContextInText] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable token and hierarchy settings for Get-OfficeDocumentHierarchy.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeReaderHierarchyOptions -MaxTokens 500 -OverlapTokens 50 -IncludeContextInText
Get-OfficeDocumentHierarchy -Path .\handbook.pdf -ChunkingOptions $options
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

### -MaxContextCharacters
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

### -MaxHierarchyDepth
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

### -MaxInputChunks
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

### -MaxOutputChunks
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

### -MaxTokens
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
