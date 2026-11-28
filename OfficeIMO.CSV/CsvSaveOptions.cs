using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.CSV;

/// <summary>
/// Options controlling how CSV content is serialized.
/// </summary>
public sealed class CsvSaveOptions
{
    private const char DefaultDelimiter = ',';
    internal CsvFormulaInjectionPolicy? ConfiguredFormulaInjectionPolicy { get; private set; }

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get; set; } = DefaultDelimiter;

    /// <summary>
    /// Gets or sets the field delimiter text. Leave unset to use <see cref="Delimiter"/>.
    /// Single-character values keep the optimized character delimiter path; longer values enable flexible delimiter writing.
    /// </summary>
    public string? DelimiterText { get; set; }

    /// <summary>
    /// Gets or sets the newline sequence written between records. Default is <see cref="Environment.NewLine"/>.
    /// </summary>
    public string NewLine { get; set; } = Environment.NewLine;

    /// <summary>
    /// Gets or sets whether to include the header row when writing. Default is <c>true</c>.
    /// </summary>
    public bool IncludeHeader { get; set; } = true;

    /// <summary>
    /// Gets or sets the culture used for formatting values. Defaults to <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Gets or sets the text encoding used when writing to files or streams. Defaults to UTF-8 without BOM when omitted.
    /// </summary>
    public Encoding? Encoding { get; set; }

    /// <summary>
    /// Gets or sets compression used when writing files. Default infers compression from the file extension.
    /// </summary>
    public CsvCompressionType CompressionType { get; set; } = CsvCompressionType.Auto;

    /// <summary>
    /// Gets or sets the compression level used when writing compressed CSV files.
    /// </summary>
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;

    /// <summary>
    /// Gets or sets the token written for <c>null</c> values. Defaults to an empty field.
    /// </summary>
    public string? NullValue { get; set; }

    /// <summary>
    /// Gets or sets a custom date/time format used when writing <see cref="DateTime"/> and <see cref="DateTimeOffset"/> values.
    /// </summary>
    public string? DateTimeFormat { get; set; }

    /// <summary>
    /// Gets or sets whether date/time values are converted to UTC before formatting.
    /// </summary>
    public bool UseUtc { get; set; }

    /// <summary>
    /// Gets or sets whether file output appends to an existing CSV file.
    /// </summary>
    public bool Append { get; set; }

    /// <summary>
    /// Gets or sets whether file output should fail if the destination already exists.
    /// </summary>
    public bool NoClobber { get; set; }

    /// <summary>
    /// Gets or sets how formula-like values are handled before writing CSV output. Default preserves values exactly.
    /// </summary>
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy
    {
        get => ConfiguredFormulaInjectionPolicy ?? CsvFormulaInjectionPolicy.Preserve;
        set => ConfiguredFormulaInjectionPolicy = value;
    }

    /// <summary>
    /// Gets or sets when fields are quoted. Default quotes only fields that need quoting.
    /// </summary>
    public CsvQuoteMode QuoteMode { get; set; } = CsvQuoteMode.AsNeeded;

    /// <summary>
    /// Gets or sets escape encoding for fields. Default is <see cref="CsvEscapeMode.None"/>.
    /// <see cref="CsvEscapeMode.Backslash"/> requires <see cref="QuoteMode"/> to be <see cref="CsvQuoteMode.Never"/>.
    /// </summary>
    public CsvEscapeMode EscapeMode { get; set; }

    /// <summary>
    /// PowerShell <c>ConvertTo-Csv -UseQuotes Never</c> parity for PSWriteOffice: unquoted fields are
    /// written unchanged even when they contain the delimiter or a line break.
    /// </summary>
    internal bool WriteUnquotedFieldsLiterally { get; set; }

    // Writers carry one quote-mode value; literal and backslash-escaped output are internal unquoted variants.
    internal CsvQuoteMode ResolvedQuoteMode => EscapeMode switch
    {
        CsvEscapeMode.None when QuoteMode == CsvQuoteMode.Never && WriteUnquotedFieldsLiterally => CsvWriter.LiteralUnquotedQuoteMode,
        CsvEscapeMode.None => QuoteMode,
        CsvEscapeMode.Backslash when QuoteMode == CsvQuoteMode.Never => CsvWriter.BackslashEscapedQuoteMode,
        CsvEscapeMode.Backslash => throw new ArgumentException(
            "Backslash escapes require QuoteMode.Never; quoted fields already protect delimiters and line breaks."),
        _ => throw new ArgumentOutOfRangeException(nameof(EscapeMode))
    };

    /// <summary>
    /// Gets or sets the character that encloses quoted fields. Default is <c>"</c>. Fields that contain it are quoted.
    /// </summary>
    public char QuoteCharacter { get; set; } = '"';

    /// <summary>
    /// Gets or sets the character written before the quote character, and before itself, inside quoted fields, such
    /// as <c>\</c>. Leave unset to double the quote character as RFC 4180 does.
    /// </summary>
    public char? EscapeCharacter { get; set; }

    /// <summary>Gets whether fields use a quote or escape character other than the RFC 4180 double quote.</summary>
    internal bool UsesCustomQuoting =>
        (QuoteCharacter != '"' || (EscapeCharacter is not null && EscapeCharacter != '"'))
        && ResolvedQuoteMode is CsvQuoteMode.AsNeeded or CsvQuoteMode.Always;

    /// <summary>
    /// Gets or sets field names that should always be quoted when <see cref="QuoteMode"/> is <see cref="CsvQuoteMode.AsNeeded"/>.
    /// </summary>
    public string[]? QuoteFields { get; set; }

    /// <summary>
    /// Creates an independent copy of these save options.
    /// </summary>
    public CsvSaveOptions Clone()
    {
        var clone = (CsvSaveOptions)MemberwiseClone();
        clone.QuoteFields = QuoteFields is null ? null : (string[])QuoteFields.Clone();
        return clone;
    }
}
