---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentStructured
## SYNOPSIS
Extracts a bounded schema-friendly view of a supported document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentStructured [-Path] <string> [-ReaderOptions <ReaderOptions>] [-ExtractionOptions <OfficeDocumentStructuredExtractionOptions>] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Extracts a bounded schema-friendly view of a supported document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Get-OfficeDocumentStructured -Path .\report.docx; $result.Records | Group-Object Category
```

Returns deterministic structured records and source diagnostics without format-specific parsing.

## PARAMETERS

### -ExtractionOptions
Optional structured extraction categories and limits.

```yaml
Type: OfficeDocumentStructuredExtractionOptions
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

- `OfficeIMO.Reader.OfficeDocumentStructuredExtractionResult`

## RELATED LINKS

- None
