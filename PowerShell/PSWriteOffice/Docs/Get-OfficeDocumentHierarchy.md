---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentHierarchy
## SYNOPSIS
Creates bounded token-aware chunks and a deterministic document hierarchy.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentHierarchy [-Path] <string> [-ReaderOptions <ReaderOptions>] [-ChunkingOptions <ReaderHierarchicalChunkingOptions>] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Creates bounded token-aware chunks and a deterministic document hierarchy.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeReaderHierarchyOptions -MaxTokens 500 -OverlapTokens 50 -IncludeContextInText
$result = Get-OfficeDocumentHierarchy -Path .\handbook.pdf -ChunkingOptions $options
```

Returns chunks, token evidence, overlap counts, and flattened parent/child nodes.

## PARAMETERS

### -ChunkingOptions
Optional token budget, overlap, hierarchy, and token-counter settings.

```yaml
Type: ReaderHierarchicalChunkingOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Reader
Optional immutable OfficeIMO reader with caller-configured handlers and processors.

```yaml
Type: OfficeDocumentReader
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReaderOptions
Optional source-reading limits and format behavior.

```yaml
Type: ReaderOptions
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Reader.ReaderChunkHierarchyResult`

## RELATED LINKS

- None
