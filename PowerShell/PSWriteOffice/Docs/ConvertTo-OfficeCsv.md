---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeCsv
## SYNOPSIS
Converts objects or a CSV document into CSV text.

## SYNTAX
### InputObjectDelimiter (Default)
```powershell
ConvertTo-OfficeCsv [-InputObject <Object>] [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [<CommonParameters>]
```

### DocumentDelimiter
```powershell
ConvertTo-OfficeCsv -Document <CsvDocument> [-Delimiter <char>] [-DelimiterText <string>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [<CommonParameters>]
```

### DocumentCulture
```powershell
ConvertTo-OfficeCsv -Document <CsvDocument> -UseCulture [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [<CommonParameters>]
```

### InputObjectCulture
```powershell
ConvertTo-OfficeCsv -UseCulture [-InputObject <Object>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-NullValue <string>] [-DateTimeFormat <string>] [-UseUtc] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [<CommonParameters>]
```

## DESCRIPTION
Use Export-OfficeCsv when the destination is a file.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $csv = $data | ConvertTo-OfficeCsv
```

Generates CSV text from the input objects.

### EXAMPLE 2
```powershell
PS> $rows = @(
  [ordered]@{ Id = 1; Name = 'Alpha'; Total = 10.5 },
  [ordered]@{ Id = 2; Name = 'Beta'; Total = 7.25 }
)
$csv = $rows | ConvertTo-OfficeCsv -Delimiter ';'
```

Uses ordered dictionaries to enforce column order and a custom delimiter.

### EXAMPLE 3
```powershell
PS> $csv = $data | ConvertTo-OfficeCsv -NoHeader
```

Writes rows only when a downstream system expects headerless CSV.

## PARAMETERS

### -Culture
Culture used for value formatting.

```yaml
Type: CultureInfo
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
CSV document to serialize.

```yaml
Type: CsvDocument
Parameter Sets: DocumentDelimiter, DocumentCulture
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FormulaInjectionPolicy
Controls how formula-like values are written.

```yaml
Type: CsvFormulaInjectionPolicy
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
Aliases: None
Possible values: Preserve, Escape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects to convert into CSV rows.

```yaml
Type: Object
Parameter Sets: InputObjectDelimiter, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary field. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: InputObjectDelimiter, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
Aliases: None
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -QuoteFields
Field names that should always be quoted when UseQuotes is AsNeeded.

```yaml
Type: String[]
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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
Parameter Sets: DocumentCulture, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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
Parameter Sets: InputObjectDelimiter, DocumentDelimiter, DocumentCulture, InputObjectCulture
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

- `OfficeIMO.CSV.CsvDocument`
- `System.Object`

## OUTPUTS

- `System.String`

## RELATED LINKS

- None
