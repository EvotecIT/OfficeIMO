---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeDocumentReaderOptions
## SYNOPSIS
Creates common read, detection, package and folder policies for document readers.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeDocumentReaderOptions [-Projection <ReaderProjection>] [-DetectionMode <ReaderDetectionMode>] [-MaximumDetectionProbeBytes <Int32>] [-MaximumDetectionContainerEntries <Int32>] [-MaximumPartUncompressedBytes <Int64>] [-MaximumTotalUncompressedBytes <Int64>] [-MaximumImageAssets <Int32>] [-MaximumImageAssetBytes <Int64>] [-MaximumTotalImageAssetBytes <Int64>] [-MaximumImagePlacementsPerRelationship <Int32>] [-MaximumTraversalEntries <Int32>] [-FollowSymlink] [-NoDeterministicOrder] [-Encoding <Encoding>] [-FailOnInvalidTextBytes] [<CommonParameters>]
```

## DESCRIPTION
Creates common read, detection, package and folder policies for document readers.

## EXAMPLES

### EXAMPLE 1
```powershell
Read-OfficeDocument ./Report.docx -As Text -Options (New-OfficeDocumentReaderOptions -Projection TextAndMetadata)
```


### EXAMPLE 2
```powershell
$options = New-OfficeDocumentReaderOptions -DetectionMode PreferContent -MaximumTraversalEntries 50000
$reader = New-OfficeDocumentReader -Options $options
Read-OfficeDocument ./Reports -Recurse -Reader $reader
```


## PARAMETERS

### -DetectionMode
How extension and content evidence select the format handler.

```yaml
Type: ReaderDetectionMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: ExtensionOnly, ContentWhenUnknown, PreferContent

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Encoding
Text encoding for sources without a byte-order mark.

```yaml
Type: Encoding
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnInvalidTextBytes
Reject invalid byte sequences instead of replacing them.

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

### -FollowSymlink
Follow links and junctions during folder traversal. Native format admission rules still apply.

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

### -MaximumDetectionContainerEntries
Maximum archive entries inspected during format detection.

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

### -MaximumDetectionProbeBytes
Maximum prefix bytes inspected during content detection.

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

### -MaximumImageAssetBytes
Maximum bytes retained for one Office image asset.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumImageAssets
Maximum retained Office image assets; zero retains none.

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

### -MaximumImagePlacementsPerRelationship
Maximum placements retained for one Office image relationship.

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

### -MaximumPartUncompressedBytes
Maximum uncompressed bytes in one Office package part.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalImageAssetBytes
Maximum combined retained Office image payload bytes.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalUncompressedBytes
Maximum combined uncompressed bytes in an Office package.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTraversalEntries
Maximum raw file-system entries inspected before filtering.

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

### -NoDeterministicOrder
Do not sort folder traversal results.

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

### -Projection
Derived representations to retain.

```yaml
Type: ReaderProjection
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Full, TextAndMetadata

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

- `OfficeIMO.Reader.ReaderOptions`

## RELATED LINKS

- None
