---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeOpenDocument
## SYNOPSIS
Loads a native ODT, ODS, ODP, or ODG document.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeOpenDocument [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-Show] [-Password <securestring>] [-MaxPackageBytes <Int64>] [-MaxEntries <Int32>] [-MaxEntryUncompressedBytes <Int64>] [-MaxTotalUncompressedBytes <Int64>] [-MaxTotalKdfIterations <Int64>] [-MaxCompressionRatio <Double>] [-MaxDepth <Int32>] [-MaxXmlCharacters <Int64>] [-MaxXmlDepth <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeOpenDocument [[-Content] <scriptblock>] -LiteralPath <string> [-PassThru] [-Show] [-Password <securestring>] [-MaxPackageBytes <Int64>] [-MaxEntries <Int32>] [-MaxEntryUncompressedBytes <Int64>] [-MaxTotalUncompressedBytes <Int64>] [-MaxTotalKdfIterations <Int64>] [-MaxCompressionRatio <Double>] [-MaxDepth <Int32>] [-MaxXmlCharacters <Int64>] [-MaxXmlDepth <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Loads a native ODT, ODS, ODP, or ODG document.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeOpenDocument -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Open-OfficeOpenDocument -LiteralPath 'C:\Path'
```


## PARAMETERS

### -Content
An optional authoring block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact source path without wildcard expansion.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaxCompressionRatio
Maximum declared expansion ratio for a compressed entry.

```yaml
Type: Double
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the FileInfo after a completed authoring block.

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

### -Password
Password used to decrypt an encrypted OpenDocument package.

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
Path to an ODT, ODS, ODP, or ODG file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Show
Open the saved document after a completed authoring block.

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
