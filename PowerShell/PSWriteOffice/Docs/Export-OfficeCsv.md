---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeCsv
## SYNOPSIS
Exports objects or a CSV document to a CSV file.

## SYNTAX
### InputObjectPathDelimiter (Default)
```powershell
Export-OfficeCsv [-Path] <string> [-InputObject <Object>] [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectPathCulture
```powershell
Export-OfficeCsv [-Path] <string> -UseCulture [-InputObject <Object>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectLiteralPathDelimiter
```powershell
Export-OfficeCsv -LiteralPath <string> [-InputObject <Object>] [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectLiteralPathCulture
```powershell
Export-OfficeCsv -LiteralPath <string> -UseCulture [-InputObject <Object>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentPathDelimiter
```powershell
Export-OfficeCsv [-Path] <string> -Document <CsvDocument> [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentPathCulture
```powershell
Export-OfficeCsv [-Path] <string> -Document <CsvDocument> -UseCulture [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentLiteralPathDelimiter
```powershell
Export-OfficeCsv -Document <CsvDocument> -LiteralPath <string> [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentLiteralPathCulture
```powershell
Export-OfficeCsv -Document <CsvDocument> -LiteralPath <string> -UseCulture [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-CompressionType <CsvCompressionType>] [-CompressionLevel <CompressionLevel>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use ConvertTo-OfficeCsv when the destination should be CSV text in the pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $data | Export-OfficeCsv -Path .\export.csv
```

Streams PowerShell objects into a CSV file.

### EXAMPLE 2
```powershell
PS> $data | Export-OfficeCsv -Path .\export.csv -UseCulture -Culture pl-PL
```

Uses the selected culture list separator as the delimiter.

## PARAMETERS

### -Append
Append rows to an existing CSV file. Existing headers are reused when present.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CompressionLevel
Compression level used when CompressionType is enabled.

```yaml
Type: CompressionLevel
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: Optimal, Fastest, NoCompression, SmallestSize

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CompressionType
Compression used when writing the CSV file.

```yaml
Type: CsvCompressionType
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: None, Auto, GZip, Deflate, Brotli, ZLib

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Culture
Culture used for value formatting.

```yaml
Type: CultureInfo
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateTimeFormat
Date/time format used for DateTime and DateTimeOffset values.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delimiter
Field delimiter character.

```yaml
Type: Char
Parameter Sets: InputObjectPathDelimiter, InputObjectLiteralPathDelimiter, DocumentPathDelimiter, DocumentLiteralPathDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DelimiterText
Field delimiter text for multi-character delimiters such as || or ::.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectLiteralPathDelimiter, DocumentPathDelimiter, DocumentLiteralPathDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
CSV document to export.

```yaml
Type: CsvDocument
Parameter Sets: DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Encoding
Encoding used when writing files.

```yaml
Type: Encoding
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow overwriting read-only files and appending rows with missing existing columns.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FormulaInjectionPolicy
Controls how formula-like values are written.

```yaml
Type: CsvFormulaInjectionPolicy
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: Preserve, Escape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects to export into CSV rows.

```yaml
Type: Object
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Literal destination CSV path. Wildcards are not expanded.

```yaml
Type: String
Parameter Sets: InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: PSPath, LP
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary field. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one field value. Defaults to 64; increase explicitly for trusted deeper values.

```yaml
Type: Int32
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewLine
Override the newline sequence.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Do not overwrite an existing CSV file.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: NoOverwrite
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Omit the header row from the output.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NullValue
Token written for null values.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a FileInfo for the exported file.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination CSV path.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, DocumentPathDelimiter, DocumentPathCulture
Aliases: FilePath, OutputPath, OutPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -QuoteFields
Field names that should always be quoted when UseQuotes is AsNeeded.

```yaml
Type: String[]
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseCulture
Use the list separator from the selected or current culture as the delimiter.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathCulture, InputObjectLiteralPathCulture, DocumentPathCulture, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseQuotes
Controls when CSV fields are quoted. Defaults to quoting only fields that need it.

```yaml
Type: CsvQuoteMode
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: AsNeeded, Always, Never

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseUtc
Convert date/time values to UTC before formatting.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
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
- `OfficeIMO.CSV.CsvDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
