---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeCsv
## SYNOPSIS
Loads a CSV document from disk or parses CSV text.

## SYNTAX
### PathDelimiter (Default)
```powershell
Get-OfficeCsv [-Path] <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### PathCulture
```powershell
Get-OfficeCsv [-Path] <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### PathDetect
```powershell
Get-OfficeCsv [-Path] <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### LiteralPathDelimiter
```powershell
Get-OfficeCsv -LiteralPath <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### LiteralPathCulture
```powershell
Get-OfficeCsv -LiteralPath <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### LiteralPathDetect
```powershell
Get-OfficeCsv -LiteralPath <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### TextDelimiter
```powershell
Get-OfficeCsv -Text <string> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### TextCulture
```powershell
Get-OfficeCsv -Text <string> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

### TextDetect
```powershell
Get-OfficeCsv -Text <string> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Returns an CsvDocument for inspection or further transformations.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $csv = Get-OfficeCsv -Path .\data.csv
```

Loads the CSV file into an OfficeIMO CsvDocument.

### EXAMPLE 2
```powershell
PS> $csv = Get-OfficeCsv -Text \"Name;Total`nAlpha;10\" -Delimiter ';'
```

Parses a semicolon-delimited CSV string into a document.

### EXAMPLE 3
```powershell
PS> $csv = Get-OfficeCsv -Path .\data.csv; $csv.Header
```

Returns the header list so you can verify the expected column names.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectParseErrors
Collect parse errors and write them as non-terminating errors after each input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnCountMismatchPolicy
Controls how rows with fewer or more fields than the header are handled.

```yaml
Type: CsvColumnCountMismatchPolicy
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Strict, PadMissingFieldsAndIgnoreExtraFields

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CommentCharacter
Character that identifies comment rows.

```yaml
Type: Char
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CompressionType
Compression used when reading files. Auto infers from the file extension.

```yaml
Type: CsvCompressionType
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values: None, Auto, GZip, Deflate, Brotli, ZLib

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Culture
Culture used for type conversions.

```yaml
Type: CultureInfo
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateTimeFormats
Additional date/time formats used by typed conversions and validation.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter, TextDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DelimiterCandidates
Delimiter candidates to consider when detecting the delimiter.

```yaml
Type: Char[]
Parameter Sets: PathDetect, LiteralPathDetect, TextDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter, TextDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DetectDelimiter
Detect the delimiter from the first meaningful records.

```yaml
Type: SwitchParameter
Parameter Sets: PathDetect, LiteralPathDetect, TextDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DuplicateHeaderBehavior
Controls how duplicate header names are handled.

```yaml
Type: CsvDuplicateHeaderBehavior
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Preserve, Rename, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Encoding
Encoding used when reading the file.

```yaml
Type: Encoding
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Explicit header names to use; when provided, the first CSV record is treated as data.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InternStrings
Reuse repeated string values through a per-read cache.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal path to one or more CSV files.

```yaml
Type: String[]
Parameter Sets: LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: PSPath, LP
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaxDecompressedBytes
Maximum decompressed bytes to read from compressed CSV files.

```yaml
Type: Int64
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxFieldLength
Maximum length allowed for any parsed field.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxParseErrors
Maximum number of collected parse errors before parsing fails.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxQuotedFieldLength
Maximum length allowed for fields parsed from quoted records.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Treat the first record as data and generate default column names.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NormalizeQuotes
Normalize curly quote characters to straight quotes.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NullValue
Token that is materialized as null when loading rows.

```yaml
Type: String
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ParseErrorAction
How parse errors are handled.

```yaml
Type: CsvParseErrorAction
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Throw, SkipRow

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to one or more CSV files. Wildcards are supported.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect
Aliases: FilePath, InputPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -ProgressInterval
Report progress every N parsed records.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -QuoteParsingMode
Controls whether malformed quoted fields are parsed leniently or rejected.

```yaml
Type: CsvQuoteParsingMode
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Lenient, Strict

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RecognizeW3CFieldsHeader
Recognize W3C Extended Log File Format #Fields: rows as headers.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCommentRows
Skip comment rows throughout the file.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCommentRowsBeforeHeader
Skip comment rows starting with # while discovering the header.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipRows
Number of parsed CSV records to skip before header discovery or data output.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StaticColumns
Static columns appended to every loaded row.

```yaml
Type: IDictionary
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
CSV text to parse.

```yaml
Type: String
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathCulture, LiteralPathCulture, TextCulture
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.CSV.CsvDocument`

## RELATED LINKS

- None
