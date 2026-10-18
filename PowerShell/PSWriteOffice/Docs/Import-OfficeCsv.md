---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeCsv
## SYNOPSIS
Imports CSV rows as PSCustomObjects, dictionaries, or a DataTable.

## SYNTAX
### PathDelimiter (Default)
```powershell
Import-OfficeCsv [-Path] <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### Document
```powershell
Import-OfficeCsv [-Document <CsvDocument>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### PathCulture
```powershell
Import-OfficeCsv [-Path] <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### PathDetect
```powershell
Import-OfficeCsv [-Path] <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### LiteralPathDelimiter
```powershell
Import-OfficeCsv -LiteralPath <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### LiteralPathCulture
```powershell
Import-OfficeCsv -LiteralPath <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

### LiteralPathDetect
```powershell
Import-OfficeCsv -LiteralPath <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <Int64>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <Int32>] [-InferSchema] [-SchemaSampleSize <int>] [-ColumnType <IDictionary>] [-AsHashtable] [-AsDataReader] [-AsDataTable] [<CommonParameters>]
```

## DESCRIPTION
Uses the CSV header to map fields to property names.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Import-OfficeCsv -Path .\data.csv | Format-Table
```

Imports each row as a PSCustomObject.

### EXAMPLE 2
```powershell
PS> Import-OfficeCsv -Path .\data.csv -AsHashtable | ForEach-Object { $_['Name'] }
```

Uses hashtables for dynamic schemas or key-based access.

### EXAMPLE 3
```powershell
PS> Import-OfficeCsv -Path .\data.csv -AsDataTable
```

Emits one DataTable per input file for database and table-oriented workflows.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsDataReader
Emit a forward-only IDataReader for database bulk-copy workflows.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsDataTable
Emit one DataTable per input file instead of enumerating row objects.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsHashtable
Emit dictionaries instead of PSCustomObjects.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectParseErrors
Collect parse errors and write them as non-terminating errors after each file.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values: Strict, PadMissingFieldsAndIgnoreExtraFields

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnType
Explicit column types used when emitting DataTable or IDataReader output.

```yaml
Type: IDictionary
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter
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
Parameter Sets: PathDetect, LiteralPathDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter
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
Parameter Sets: PathDetect, LiteralPathDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
CSV document to read when already loaded.

```yaml
Type: CsvDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DuplicateHeaderBehavior
Controls how duplicate header names are handled.

```yaml
Type: CsvDuplicateHeaderBehavior
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InferSchema
Infer typed columns when -AsDataTable or -AsDataReader is used.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mode
Load mode controlling materialization.

```yaml
Type: CsvReadMode
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values: Stream, InMemory

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NullValue
Token that is materialized as null when importing rows.

```yaml
Type: String
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Aliases: FilePath
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SchemaSampleSize
Maximum row count inspected when schema inference is enabled.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StaticColumns
Static columns appended to every imported row.

```yaml
Type: IDictionary
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathCulture, LiteralPathCulture
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

- `OfficeIMO.CSV.CsvDocument`
- `System.String[]`

## OUTPUTS

- `None`

## RELATED LINKS

- None
