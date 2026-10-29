---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeDocumentInfo
## SYNOPSIS
Inspects format, native counts and encryption without changing a source or reopening a live document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeDocumentInfo [-Path] <string[]> [-Detailed] [-Options <Object[]>] [-Password <securestring>] [-DetectionMode <ReaderDetectionMode>] [-MaximumProbeBytes <Int32>] [-MaximumContainerEntries <Int32>] [-NoContainerInspection] [-Reader <OfficeDocumentReader>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeDocumentInfo [-Document] <Object> [-Detailed] [-Options <Object[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeDocumentInfo -LiteralPath <string[]> [-Detailed] [-Options <Object[]>] [-Password <securestring>] [-DetectionMode <ReaderDetectionMode>] [-MaximumProbeBytes <Int32>] [-MaximumContainerEntries <Int32>] [-NoContainerInspection] [-Reader <OfficeDocumentReader>] [-ThrottleLimit <int>] [<CommonParameters>]
```

## DESCRIPTION
Inspects format, native counts and encryption without changing a source or reopening a live document.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem ./Reports/*.xlsx | Get-OfficeDocumentInfo | Select-Object Path, Format, SheetCount, IsEncrypted
```


### EXAMPLE 2
```powershell
$presentation | Get-OfficeDocumentInfo -Detailed
```


## PARAMETERS

### -Detailed
Include the native format's detached inspection snapshot.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DetectionMode
Content and filename selection policy.

```yaml
Type: ReaderDetectionMode
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: ExtensionOnly, ContentWhenUnknown, PreferContent

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live native model or normalized Reader result. Its source path is never reopened.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaximumContainerEntries
Maximum container entries examined during format detection.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumProbeBytes
Maximum prefix bytes examined for content signatures.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoContainerInspection
Skip structural container inspection when identifying the format.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Inspection, native format or Reader policy objects; hashtables are accepted.

```yaml
Type: Object[]
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used locally to open an encrypted source.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets input paths with provider wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Reader
Immutable caller-configured format registry.

```yaml
Type: OfficeDocumentReader
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrottleLimit
Gets or sets the maximum number of concurrent background operations.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

- `System.Object`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.OfficeDocumentInfo`

## RELATED LINKS

- None
