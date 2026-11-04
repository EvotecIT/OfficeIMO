---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeDocumentAsset
## SYNOPSIS
Writes document media, tables or visuals using the native Reader materializers.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeDocumentAsset [-Path] <string[]> [-OutputDirectory] <string> [-Type <OfficeDocumentAssetType[]>] [-Kind <string[]>] [-MediaType <string[]>] [-Extension <string[]>] [-TableFormat <OfficeTableExportFormat[]>] [-NoPayload] [-NoMetadata] [-Indented] [-ValidatePayloadHash] [-NoClobber] [-Sheet <OfficeSheetSelector>] [-Range <string>] [-MaximumInputBytes <Int64>] [-Reader <OfficeDocumentReader>] [-Options <Object[]>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficeDocumentAsset [-OutputDirectory] <string> -LiteralPath <string[]> [-Type <OfficeDocumentAssetType[]>] [-Kind <string[]>] [-MediaType <string[]>] [-Extension <string[]>] [-TableFormat <OfficeTableExportFormat[]>] [-NoPayload] [-NoMetadata] [-Indented] [-ValidatePayloadHash] [-NoClobber] [-Sheet <OfficeSheetSelector>] [-Range <string>] [-MaximumInputBytes <Int64>] [-Reader <OfficeDocumentReader>] [-Options <Object[]>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Writes document media, tables or visuals using the native Reader materializers.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeDocumentAsset ./Review.pptx ./Assets
```


### EXAMPLE 2
```powershell
Get-ChildItem ./Reports/*.docx | Export-OfficeDocumentAsset -OutputDirectory ./Tables -Type Table -TableFormat Csv
```


### EXAMPLE 3
```powershell
Export-OfficeDocumentAsset ./Brochure.docx ./Images -Extension png,jpg -NoClobber
```


## PARAMETERS

### -Extension
Selected media extensions, with or without a leading dot.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Indented
Indent table and visual JSON.

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

### -Kind
Selected media asset kinds.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
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

### -MaximumInputBytes
Maximum input bytes per source.

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

### -MediaType
Selected media MIME types.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Leave existing output files unchanged and return skipped materialization results.

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

### -NoMetadata
Write visual payloads without JSON metadata.

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

### -NoPayload
Write only visual metadata.

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
Typed common and format read policies, or common/format-keyed hashtables.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder for the native materialized files.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: 1
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

### -Range
A1 range within the selected worksheet.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reader
Caller-configured immutable Reader.

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

### -Sheet
Worksheet name, one-based position or sheet object.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableFormat
Table representations to write. All formats are included by default.

```yaml
Type: OfficeTableExportFormat[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: Csv, Markdown, Json

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

### -Type
Media, tables or visuals. Media is the default.

```yaml
Type: OfficeDocumentAssetType[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: Media, Table, Visual

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidatePayloadHash
Validate media payload hashes when the source supplies them.

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

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentMaterializedAsset`
- `OfficeIMO.Reader.ReaderTableMaterializedExport`
- `OfficeIMO.Reader.ReaderVisualMaterializedExport`

## RELATED LINKS

- None
