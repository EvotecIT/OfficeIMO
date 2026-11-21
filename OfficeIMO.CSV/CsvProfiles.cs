using System.Text;

namespace OfficeIMO.CSV;

/// <summary>Identifies a CSV policy preset. Presets do not change the default options.</summary>
public enum CsvProfile
{
    /// <summary>Reject malformed quotes and field-count mismatches, with fixed UTF-8 decoding.</summary>
    Strict,
    /// <summary>Preserve the existing forgiving import and lossless serialization defaults.</summary>
    Compatibility,
    /// <summary>Read spreadsheet exports and escape formula-like text when writing spreadsheet input.</summary>
    Spreadsheet,
    /// <summary>
    /// IANA <c>text/tab-separated-values</c>: tab delimiter, no quoting and one record per line.
    /// Quote characters are data; fields cannot contain tabs or line breaks unless
    /// <see cref="CsvEscapeMode.Backslash"/> is selected on the load and save options.
    /// </summary>
    Tsv
}

/// <summary>Creates independent, editable options for common CSV workflows.</summary>
public static class CsvProfiles
{
    /// <summary>
    /// Creates load options. Strict uses fixed UTF-8 with invalid-byte rejection, literal comment
    /// rows, strict quoting and column counts, duplicate-header rejection and redacted mapping
    /// failures. Spreadsheet detects delimiters and BOM encodings, keeps strict quoting and
    /// counts, and repairs missing or duplicate headers. Tsv reads tab-separated records with quotes
    /// as data, strict column counts, BOM-detected encodings and repaired headers. Compatibility
    /// retains existing defaults.
    /// </summary>
    public static CsvLoadOptions CreateLoadOptions(CsvProfile profile) =>
        new CsvLoadOptions { Parsing = CreateParsingSettings(profile) };

    /// <summary>Creates forward-only reader options with the same parsing policy as the selected load profile.</summary>
    public static CsvReaderOptions CreateReaderOptions(CsvProfile profile) =>
        new CsvReaderOptions { Parsing = CreateParsingSettings(profile) };

    private static CsvParsingSettings CreateParsingSettings(CsvProfile profile)
    {
        Validate(profile);
        if (profile == CsvProfile.Compatibility) return new CsvParsingSettings();
        if (profile == CsvProfile.Tsv)
        {
            return new CsvParsingSettings
            {
                Delimiter = '\t',
                QuoteParsingMode = CsvQuoteParsingMode.None,
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict,
                DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Rename,
                // TSV has no comment syntax; a leading '#' is data.
                SkipCommentRowsBeforeHeader = false,
                RecognizeW3CFieldsHeader = false,
                MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact,
                // Spreadsheet "Unicode text" exports are UTF-16 with a BOM.
                Encoding = new UTF8Encoding(false)
            };
        }

        return new CsvParsingSettings
        {
            QuoteParsingMode = CsvQuoteParsingMode.Strict,
            ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict,
            DuplicateHeaderBehavior = profile == CsvProfile.Strict
                ? CsvDuplicateHeaderBehavior.Throw : CsvDuplicateHeaderBehavior.Rename,
            GenerateMissingHeaderNames = profile != CsvProfile.Strict,
            SkipCommentRowsBeforeHeader = false,
            RecognizeW3CFieldsHeader = false,
            MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact,
            DetectDelimiter = profile == CsvProfile.Spreadsheet,
            DetectEncodingFromByteOrderMarks = profile == CsvProfile.Spreadsheet,
            // A matching UTF-8 preamble is accepted without permitting a BOM to replace
            // the strict decoder with a different encoding or replacement fallback.
            Encoding = profile == CsvProfile.Strict ? new UTF8Encoding(true, true) : new UTF8Encoding(false)
        };
    }

    /// <summary>
    /// Creates save options. Strict writes UTF-8 without BOM and CRLF while preserving values.
    /// Spreadsheet writes UTF-8 with BOM and CRLF and prefixes formula-like text with an
    /// apostrophe. Tsv writes tab-separated UTF-8 without BOM, LF line endings and no quoting;
    /// a field containing a tab or line break fails the write unless backslash escaping is selected.
    /// Compatibility retains existing defaults. TextWriter and string output
    /// have no byte encoding; the encoding applies to file and stream output.
    /// </summary>
    public static CsvSaveOptions CreateSaveOptions(CsvProfile profile)
    {
        Validate(profile);
        if (profile == CsvProfile.Compatibility) return new CsvSaveOptions();
        if (profile == CsvProfile.Tsv)
        {
            return new CsvSaveOptions
            {
                Delimiter = '\t',
                QuoteMode = CsvQuoteMode.Never,
                Encoding = new UTF8Encoding(false, true),
                NewLine = "\n"
            };
        }

        return new CsvSaveOptions
        {
            Encoding = new UTF8Encoding(profile == CsvProfile.Spreadsheet, true),
            NewLine = "\r\n",
            FormulaInjectionPolicy = profile == CsvProfile.Spreadsheet
                ? CsvFormulaInjectionPolicy.Escape : CsvFormulaInjectionPolicy.Preserve
        };
    }

    // Path defaults apply only when the caller passes no options: explicit options always win.
    internal static CsvParsingSettings CreateDefaultParsingSettings(string path) =>
        IsTabSeparatedPath(path) ? CreateParsingSettings(CsvProfile.Tsv) : new CsvParsingSettings();

    internal static CsvSaveOptions CreateDefaultSaveOptions(string path) =>
        IsTabSeparatedPath(path) ? CreateSaveOptions(CsvProfile.Tsv) : new CsvSaveOptions();

    internal static CsvReaderOptions CreateDefaultReaderOptions(string path) =>
        new CsvReaderOptions { Parsing = CreateDefaultParsingSettings(path) };

    /// <summary>Returns whether a path names tab-separated values (<c>.tsv</c> or <c>.tab</c>), ignoring a compression suffix.</summary>
    internal static bool IsTabSeparatedPath(string path)
    {
        string name = CsvFile.ResolveCompression(CsvCompressionType.Auto, path) == CsvCompressionType.None
            ? path
            : Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(name);
        return string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".tab", StringComparison.OrdinalIgnoreCase);
    }

    private static void Validate(CsvProfile profile)
    {
        if (profile is not CsvProfile.Strict and not CsvProfile.Compatibility and not CsvProfile.Spreadsheet and not CsvProfile.Tsv)
            throw new ArgumentOutOfRangeException(nameof(profile));
    }
}
