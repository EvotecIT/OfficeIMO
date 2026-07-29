using System;
using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.Tabular;

/// <summary>
/// Controls the format-independent behavior of <see cref="TabularReader"/>.
/// </summary>
public sealed class TabularReadOptions {
    /// <summary>
    /// Gets or sets the workbook table to read. When omitted, workbook tables are exposed
    /// in workbook order through <see cref="TabularReader.NextResult"/>.
    /// </summary>
    public string? TableName { get; set; }

    /// <summary>Gets or sets whether the first row supplies column names.</summary>
    public bool HasHeaderRow { get; set; } = true;

    /// <summary>
    /// Gets or sets whether delimited-text column types are inferred. Spreadsheet cells
    /// retain their native value types regardless of this setting.
    /// </summary>
    public bool InferTypes { get; set; }

    /// <summary>Gets or sets the maximum rows sampled when type inference is enabled.</summary>
    public int SchemaSampleRows { get; set; } = 1000;

    /// <summary>
    /// Gets or sets a delimited-text separator. When omitted, TSV paths use tab and other
    /// delimited-text paths use comma.
    /// </summary>
    public char? Delimiter { get; set; }

    /// <summary>Gets or sets whether a delimited-text separator is detected from input.</summary>
    public bool DetectDelimiter { get; set; }

    /// <summary>Gets or sets whether unquoted delimited-text fields are trimmed.</summary>
    public bool TrimWhitespace { get; set; }

    /// <summary>Gets or sets the text encoding. The default is UTF-8.</summary>
    public Encoding? Encoding { get; set; }

    /// <summary>Gets or sets the culture used for text-to-value conversion.</summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>Gets or sets whether spreadsheet numeric values are returned as decimal where possible.</summary>
    public bool NumericAsDecimal { get; set; }

    /// <summary>Gets or sets whether spreadsheet number formats are used to identify dates.</summary>
    public bool TreatDatesUsingNumberFormat { get; set; } = true;

    /// <summary>Gets or sets whether cached spreadsheet formula results are returned.</summary>
    public bool UseCachedFormulaResult { get; set; } = true;

    /// <summary>Gets or sets the maximum accepted input size. Default: 256 MiB.</summary>
    public long MaxInputBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Gets or sets the cancellation token observed while opening and reading.</summary>
    public CancellationToken CancellationToken { get; set; }

    internal void Validate() {
        if (SchemaSampleRows <= 0) {
            throw new ArgumentOutOfRangeException(nameof(SchemaSampleRows), "Schema sample rows must be greater than zero.");
        }

        if (MaxInputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes), "Maximum input bytes must be greater than zero.");
        }
    }
}
