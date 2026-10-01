#nullable enable
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
    Spreadsheet
}

/// <summary>Creates independent, editable options for common CSV workflows.</summary>
public static class CsvProfiles
{
    /// <summary>
    /// Creates load options. Strict uses fixed UTF-8 with invalid-byte rejection, literal comment
    /// rows, strict quoting and column counts, duplicate-header rejection and redacted mapping
    /// failures. Spreadsheet detects delimiters and BOM encodings, keeps strict quoting and
    /// counts, and repairs missing or duplicate headers. Compatibility retains existing defaults.
    /// </summary>
    public static CsvLoadOptions CreateLoadOptions(CsvProfile profile)
    {
        Validate(profile);
        if (profile == CsvProfile.Compatibility) return new CsvLoadOptions();
        return new CsvLoadOptions
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
    /// apostrophe. Compatibility retains existing defaults. TextWriter and string output
    /// have no byte encoding; the encoding applies to file and stream output.
    /// </summary>
    public static CsvSaveOptions CreateSaveOptions(CsvProfile profile)
    {
        Validate(profile);
        if (profile == CsvProfile.Compatibility) return new CsvSaveOptions();
        return new CsvSaveOptions
        {
            Encoding = new UTF8Encoding(profile == CsvProfile.Spreadsheet, true),
            NewLine = "\r\n",
            FormulaInjectionPolicy = profile == CsvProfile.Spreadsheet
                ? CsvFormulaInjectionPolicy.Escape : CsvFormulaInjectionPolicy.Preserve
        };
    }

    private static void Validate(CsvProfile profile)
    {
        if (profile is not CsvProfile.Strict and not CsvProfile.Compatibility and not CsvProfile.Spreadsheet)
            throw new ArgumentOutOfRangeException(nameof(profile));
    }
}
