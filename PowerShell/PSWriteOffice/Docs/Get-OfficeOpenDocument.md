---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeOpenDocument
## SYNOPSIS
Loads a native ODT, ODS, ODP, or ODG document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeOpenDocument [-Path] <string> [-Options <OdfLoadOptions>] [-Password <string>] [-MaxPackageBytes <Int64>] [-MaxEntries <Int32>] [-MaxEntryUncompressedBytes <Int64>] [-MaxTotalUncompressedBytes <Int64>] [-MaxTotalKdfIterations <Int64>] [-MaxCompressionRatio <Double>] [-MaxDepth <Int32>] [-MaxXmlCharacters <Int64>] [-MaxXmlDepth <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Loads a native ODT, ODS, ODP, or ODG document.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeOpenDocument -Path 'C:\Path'
```


## PARAMETERS

### -MaxCompressionRatio
Maximum declared expansion ratio for a compressed entry.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxDepth
Maximum archive path depth.

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

### -MaxEntries
Maximum number of ZIP entries.

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

### -MaxEntryUncompressedBytes
Maximum uncompressed size of one package entry.

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

### -MaxPackageBytes
Maximum source package size in bytes.

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

### -MaxTotalKdfIterations
Maximum aggregate PBKDF2 iterations across encrypted entries.

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

### -MaxTotalUncompressedBytes
Maximum aggregate uncompressed package size.

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

### -MaxXmlCharacters
Maximum characters allowed in one parsed XML part.

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

### -MaxXmlDepth
Maximum element nesting depth in one parsed XML part.

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

### -Options
Optional bounded package and XML settings.

```yaml
Type: OdfLoadOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to decrypt an encrypted OpenDocument package.

```yaml
Type: String
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
Path to an ODT, ODS, ODP, or ODG file.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfDocument`
- `OfficeIMO.OpenDocument.OdtDocument`
- `OfficeIMO.OpenDocument.OdsDocument`
- `OfficeIMO.OpenDocument.OdpPresentation`
- `OfficeIMO.OpenDocument.OdgDocument`

## RELATED LINKS

- None
