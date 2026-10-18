---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeCsv
## SYNOPSIS
Converts CSV text to PSCustomObjects or dictionaries.

## SYNTAX
### TextDelimiter (Default)
```powershell
ConvertFrom-OfficeCsv [-Text] <string> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

### TextCulture
```powershell
ConvertFrom-OfficeCsv [-Text] <string> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

### TextDetect
```powershell
ConvertFrom-OfficeCsv [-Text] <string> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvReadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

## DESCRIPTION
Reads CSV text from -Text or the pipeline and maps rows by header.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = ConvertFrom-OfficeCsv -Text "Name,Value`nAlpha,1"
```

Parses CSV text and emits row objects without writing a temporary file.

### EXAMPLE 2
```powershell
PS> "Name,Value", "Alpha,1" | ConvertFrom-OfficeCsv
```

Treats piped lines as one CSV stream.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter
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
Parameter Sets: TextDetect
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
Parameter Sets: TextDelimiter
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
Parameter Sets: TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NullValue
Token that is materialized as null when converting rows.

```yaml
Type: String
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCommentRows
Skip comment rows throughout the input.

```yaml
Type: SwitchParameter
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StaticColumns
Static columns appended to every converted row.

```yaml
Type: IDictionary
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextCulture
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

- `System.String`

## OUTPUTS

- `None`

## RELATED LINKS

- None
