#nullable enable

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

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get; set; } = DefaultDelimiter;

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
    /// Gets or sets the text encoding used when writing to files. Defaults to UTF-8 without BOM when omitted.
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
    /// Gets or sets how formula-like values are handled before writing CSV output. Default preserves values exactly.
    /// </summary>
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy { get; set; } = CsvFormulaInjectionPolicy.Preserve;

    /// <summary>
    /// Gets or sets when fields are quoted. Default quotes only fields that need quoting.
    /// </summary>
    public CsvQuoteMode QuoteMode { get; set; } = CsvQuoteMode.AsNeeded;

    /// <summary>
    /// Gets or sets field names that should always be quoted when <see cref="QuoteMode"/> is <see cref="CsvQuoteMode.AsNeeded"/>.
    /// </summary>
    public string[]? QuoteFields { get; set; }
}
