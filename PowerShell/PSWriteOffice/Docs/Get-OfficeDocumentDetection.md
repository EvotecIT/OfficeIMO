---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentDetection
## SYNOPSIS
Detects a document kind from extension and bounded content evidence.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentDetection [-Path] <string> [-Mode <ReaderDetectionMode>] [-MaxProbeBytes <Int32>] [-MaxContainerEntries <Int32>] [-NoContainerInspection] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Detects a document kind from extension and bounded content evidence.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeDocumentDetection -Path .\upload.bin -Mode PreferContent
```

Returns the selected kind, confidence, media type, and evidence used by OfficeIMO.Reader.

## PARAMETERS

### -MaxContainerEntries
Maximum archive entries inspected while classifying container formats.

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

### -MaxProbeBytes
Maximum prefix bytes inspected for signatures and text markers.

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

### -Mode
Policy used to combine extension and content evidence.

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

### -NoContainerInspection
Skip structural inspection of ZIP-based containers.

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

### -Path
Path to inspect.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Reader.ReaderDetectionResult`

## RELATED LINKS

- None
