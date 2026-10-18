---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentVisual
## SYNOPSIS
Reads visual payloads discovered by OfficeIMO.Reader from a supported document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentVisual [-Path] <string> [-AsExport] [-OutputDirectory <string>] [-NoOverwrite] [-Indented] [-NoPayload] [-NoJson] [-MaxInputBytes <Int64>] [-OpenXmlMaxCharactersInPart <Int64>] [-MaxChars <Int32>] [-MaxTableRows <Int32>] [-ExcludeWordFootnotes] [-ExcludePowerPointNotes] [-NoExcelHeaders] [-ExcelChunkRows <Int32>] [-ExcelSheetName <string>] [-ExcelA1Range <string>] [-NoMarkdownHeadingChunks] [-NoHashes] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Reads visual payloads discovered by OfficeIMO.Reader from a supported document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $visuals = Get-OfficeDocumentVisual -Path .\Deck.pptx -AsExport -OutputDirectory .\reader-visuals -Indented
$visuals | Select-Object Id, ContentType, PayloadPath
```

Extracts supported visual payloads and writes payload plus metadata sidecars.

## PARAMETERS

### -AsExport
Return deterministic payload and JSON export bundles instead of visual models.

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

### -ExcelA1Range
Optional Excel A1 range to read.

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

### -ExcelChunkRows
Excel rows per emitted worksheet chunk.

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

### -ExcelSheetName
Optional Excel sheet name to read.

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

### -ExcludePowerPointNotes
Exclude PowerPoint speaker notes.

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

### -ExcludeWordFootnotes
Exclude Word footnotes.

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

### -Indented
Indent JSON payloads in export bundles and sidecars.

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

### -MaxChars
Maximum emitted chunk characters.

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

### -MaxInputBytes
Maximum input size in bytes.

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

### -MaxTableRows
Maximum table rows per emitted table.

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

### -NoExcelHeaders
Treat the first Excel row as data instead of headers.

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

### -NoHashes
Disable source and chunk hash computation.

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

### -NoJson
Do not write JSON sidecars when -OutputDirectory is used.

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

### -NoMarkdownHeadingChunks
Do not split Markdown by headings.

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

### -NoOverwrite
Do not overwrite existing sidecar files.

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

### -NoPayload
Do not write raw payload sidecars when -OutputDirectory is used.

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

### -OpenXmlMaxCharactersInPart
OpenXML maximum characters per part.

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

### -OutputDirectory
Optional directory where visual sidecars should be written.

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
File path to read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
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

- `OfficeIMO.Reader.ReaderVisual`
- `OfficeIMO.Reader.ReaderVisualExportBundle`
- `OfficeIMO.Reader.ReaderVisualMaterializedExport`

## RELATED LINKS

- None
