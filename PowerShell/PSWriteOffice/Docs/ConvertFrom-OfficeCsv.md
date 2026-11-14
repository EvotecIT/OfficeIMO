---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertFrom-OfficeCsv
## SYNOPSIS
Converts CSV text into PSCustomObjects, dictionaries, a DataTable, a data reader or a CsvDocument.

## SYNTAX
### InputObject (Default)
```powershell
ConvertFrom-OfficeCsv [-InputObject] <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <string>] [-TrimWhitespace] [-AllowEmptyLines] [-NoSkipCommentRowsBeforeHeader] [-SkipCommentRows] [-CommentCharacter <char>] [-NoW3CFieldsHeader] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-InferSchema] [-SchemaSampleRows <int>] [-ColumnType <IDictionary>] [-As <OfficeCsvOutput>] [<CommonParameters>]
```

### InputObjectDetect
```powershell
ConvertFrom-OfficeCsv [-InputObject] <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace] [-AllowEmptyLines] [-NoSkipCommentRowsBeforeHeader] [-SkipCommentRows] [-CommentCharacter <char>] [-NoW3CFieldsHeader] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-InferSchema] [-SchemaSampleRows <int>] [-ColumnType <IDictionary>] [-As <OfficeCsvOutput>] [<CommonParameters>]
```

### InputObjectCulture
```powershell
ConvertFrom-OfficeCsv [-InputObject] <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace] [-AllowEmptyLines] [-NoSkipCommentRowsBeforeHeader] [-SkipCommentRows] [-CommentCharacter <char>] [-NoW3CFieldsHeader] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <Int32>] [-MaxQuotedFieldLength <Int32>] [-NormalizeQuotes] [-InternStrings] [-InferSchema] [-SchemaSampleRows <int>] [-ColumnType <IDictionary>] [-As <OfficeCsvOutput>] [<CommonParameters>]
```

## DESCRIPTION
The reverse of ConvertTo-OfficeCsv and the text form of Import-OfficeCsv: piped strings are read as lines of one
CSV text, as ConvertFrom-Csv reads them, and parsed with Import-OfficeCsv's options and output shapes. Use
Import-OfficeCsv when the source is a file.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> "Name,Total`nAlpha,10`nBeta,7" | ConvertFrom-OfficeCsv
```

Emits one PSCustomObject per row, with property names from the header.

### EXAMPLE 2
```powershell
PS> $rows = Get-Process | Select-Object -First 5 Name, Id | ConvertTo-OfficeCsv | ConvertFrom-OfficeCsv
```

Reads the text ConvertTo-OfficeCsv writes back into objects.

### EXAMPLE 3
```powershell
PS> Get-Content ./export.txt | ConvertFrom-OfficeCsv -Delimiter ';' -As DataTable -InferSchema
```

Joins the lines Get-Content emits and infers column types from the values.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -As
Output shape: Object, Hashtable, DataTable, DataReader or Document.

```yaml
Type: OfficeCsvOutput
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values: Object, Hashtable, DataTable, DataReader, Document

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectParseErrors
Collect parse errors and write them as non-terminating errors after the rows.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Culture
Culture used for type conversions, and for the delimiter with UseCulture.

```yaml
Type: CultureInfo
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delimiter
Field delimiter, including multi-character delimiters.

```yaml
Type: String
Parameter Sets: InputObject
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
Parameter Sets: InputObjectDetect
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
Parameter Sets: InputObjectDetect
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values: Preserve, Rename, Throw

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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InferSchema
Infer typed columns when -As DataTable or -As DataReader is used.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
CSV text, as one string or as lines; piped strings are joined with line breaks.

```yaml
Type: String[]
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -InternStrings
Reuse repeated string values through a per-read cache.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSkipCommentRowsBeforeHeader
Keep comment rows while discovering the header.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoW3CFieldsHeader
Disable W3C Extended Log File Format #Fields: headers.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NullValue
Token that is materialized as null.

```yaml
Type: String
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values: Throw, SkipRow

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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values: Lenient, Strict, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SchemaSampleRows
Maximum row count inspected when schema inference is enabled.

```yaml
Type: Int32
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCommentRows
Skip comment rows throughout the text.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StaticColumns
Static columns appended to every row.

```yaml
Type: IDictionary
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Type: SwitchParameter
Parameter Sets: InputObject, InputObjectDetect, InputObjectCulture
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
Parameter Sets: InputObjectCulture
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

- `System.Management.Automation.PSObject`
- `System.Collections.Hashtable`
- `System.Data.DataTable`
- `System.Data.IDataReader`
- `OfficeIMO.CSV.CsvDocument`

## RELATED LINKS

- None
