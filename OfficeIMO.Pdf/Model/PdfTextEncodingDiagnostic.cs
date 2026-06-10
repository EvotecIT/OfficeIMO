using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes text that cannot be written through the current generated PDF text encoding path.
/// </summary>
public sealed class PdfTextEncodingDiagnostic {
    private readonly string? _code;

    /// <summary>
    /// Creates a text encoding diagnostic.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter)
        : this(source, index, codePoint, text, isControlCharacter, string.Empty, string.Empty, string.Empty) {
    }

    internal PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string? code, string? message, bool customCode) {
        Source = source ?? string.Empty;
        Index = index;
        CodePoint = codePoint ?? string.Empty;
        Text = text ?? string.Empty;
        IsControlCharacter = isControlCharacter;
        Encoding = string.Empty;
        Remediation = string.Empty;
        Location = string.Empty;
        RunIndex = null;
        PageNumber = null;
        TableRowIndex = null;
        TableColumnIndex = null;
        FieldName = string.Empty;
        _code = string.IsNullOrWhiteSpace(code) ? null : code;
        Message = string.IsNullOrWhiteSpace(message)
            ? CreateMessage(Index, CodePoint, Text, IsControlCharacter, Encoding, Remediation)
            : message!;
    }

    /// <summary>
    /// Creates a text encoding diagnostic with a custom encoding or font coverage description.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation)
        : this(source, index, codePoint, text, isControlCharacter, encoding, remediation, string.Empty) {
    }

    /// <summary>
    /// Creates a text encoding diagnostic with a custom encoding or font coverage description and document location.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation, string location)
        : this(source, index, codePoint, text, isControlCharacter, encoding, remediation, location, null) {
    }

    /// <summary>
    /// Creates a text encoding diagnostic with a custom encoding or font coverage description, document location, and rich-text run index.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <param name="runIndex">Optional zero-based rich text run index inside the generated document location.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation, string location, int? runIndex)
        : this(source, index, codePoint, text, isControlCharacter, encoding, remediation, location, runIndex, null) {
    }

    /// <summary>
    /// Creates a text encoding diagnostic with a custom encoding or font coverage description, document location, rich-text run index, and generated page number.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <param name="runIndex">Optional zero-based rich text run index inside the generated document location.</param>
    /// <param name="pageNumber">Optional one-based generated page number for page-scoped diagnostics.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation, string location, int? runIndex, int? pageNumber)
        : this(source, index, codePoint, text, isControlCharacter, encoding, remediation, location, runIndex, pageNumber, null, null) {
    }

    /// <summary>
    /// Creates a text encoding diagnostic with generated page and table-cell coordinates.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <param name="runIndex">Optional zero-based rich text run index inside the generated document location.</param>
    /// <param name="pageNumber">Optional one-based generated page number for page-scoped diagnostics.</param>
    /// <param name="tableRowIndex">Optional zero-based table row index for table-cell diagnostics.</param>
    /// <param name="tableColumnIndex">Optional zero-based table column index for table-cell diagnostics.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation, string location, int? runIndex, int? pageNumber, int? tableRowIndex, int? tableColumnIndex)
        : this(source, index, codePoint, text, isControlCharacter, encoding, remediation, location, runIndex, pageNumber, tableRowIndex, tableColumnIndex, string.Empty) {
    }

    /// <summary>
    /// Creates a text encoding diagnostic with generated page, table-cell coordinates, and form field name.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    /// <param name="encoding">Encoding or font coverage description that rejected the character.</param>
    /// <param name="remediation">Optional remediation guidance for the caller.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <param name="runIndex">Optional zero-based rich text run index inside the generated document location.</param>
    /// <param name="pageNumber">Optional one-based generated page number for page-scoped diagnostics.</param>
    /// <param name="tableRowIndex">Optional zero-based table row index for table-cell diagnostics.</param>
    /// <param name="tableColumnIndex">Optional zero-based table column index for table-cell diagnostics.</param>
    /// <param name="fieldName">Optional generated AcroForm field name for form-field diagnostics.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation, string location, int? runIndex, int? pageNumber, int? tableRowIndex, int? tableColumnIndex, string? fieldName) {
        if (pageNumber.HasValue && pageNumber.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "PDF diagnostic page number must be a positive one-based value.");
        }

        if (tableRowIndex.HasValue != tableColumnIndex.HasValue) {
            throw new ArgumentException("PDF diagnostic table row and column indexes must be supplied together.", nameof(tableRowIndex));
        }

        if (tableRowIndex.HasValue && tableRowIndex.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(tableRowIndex), tableRowIndex, "PDF diagnostic table row index must be zero or greater.");
        }

        if (tableColumnIndex.HasValue && tableColumnIndex.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(tableColumnIndex), tableColumnIndex, "PDF diagnostic table column index must be zero or greater.");
        }

        Source = source ?? string.Empty;
        Index = index;
        CodePoint = codePoint ?? string.Empty;
        Text = text ?? string.Empty;
        IsControlCharacter = isControlCharacter;
        Encoding = encoding ?? string.Empty;
        Remediation = remediation ?? string.Empty;
        Location = location ?? string.Empty;
        RunIndex = runIndex;
        PageNumber = pageNumber;
        TableRowIndex = tableRowIndex;
        TableColumnIndex = tableColumnIndex;
        FieldName = fieldName ?? string.Empty;
        _code = null;
        Message = CreateMessage(Index, CodePoint, Text, IsControlCharacter, Encoding, Remediation);
    }

    /// <summary>Caller-provided source label such as a block, field, sheet, slide, or converter area.</summary>
    public string Source { get; }

    /// <summary>UTF-16 index of the unsupported character in the inspected text.</summary>
    public int Index { get; }

    /// <summary>Unicode code point formatted as U+XXXX or U+XXXXX.</summary>
    public string CodePoint { get; }

    /// <summary>Unsupported character or surrogate pair text. Control characters are represented as an empty string.</summary>
    public string Text { get; }

    /// <summary>Whether the unsupported character is a control character.</summary>
    public bool IsControlCharacter { get; }

    /// <summary>Encoding or font coverage description that rejected the character.</summary>
    public string Encoding { get; }

    /// <summary>Optional remediation guidance for the caller.</summary>
    public string Remediation { get; }

    /// <summary>Optional generated document location such as a block, table cell, or canvas item path.</summary>
    public string Location { get; }

    /// <summary>Optional zero-based rich text run index inside the generated document location.</summary>
    public int? RunIndex { get; }

    /// <summary>Optional one-based generated page number for page-scoped diagnostics.</summary>
    public int? PageNumber { get; }

    /// <summary>Optional zero-based table row index for table-cell diagnostics.</summary>
    public int? TableRowIndex { get; }

    /// <summary>Optional zero-based table column index for table-cell diagnostics.</summary>
    public int? TableColumnIndex { get; }

    /// <summary>Optional generated AcroForm field name for form-field diagnostics.</summary>
    public string FieldName { get; }

    /// <summary>Stable warning code suitable for shared conversion reports.</summary>
    public string Code => _code ?? (IsControlCharacter ? "unsupported-control-character" : "unsupported-text-glyph");

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>
    /// Converts this text diagnostic to the shared conversion warning shape used by PDF adapters.
    /// </summary>
    /// <param name="converter">Converter or adapter name to place on the warning.</param>
    /// <returns>A shared conversion warning carrying this diagnostic and stable details.</returns>
    public PdfConversionWarning ToConversionWarning(string converter = "OfficeIMO.Pdf") {
        var details = new Dictionary<string, string> {
            ["index"] = Index.ToString(CultureInfo.InvariantCulture),
            ["codePoint"] = CodePoint,
            ["text"] = Text,
            ["isControlCharacter"] = IsControlCharacter ? "true" : "false"
        };
        if (!string.IsNullOrWhiteSpace(Encoding)) {
            details["encoding"] = Encoding;
        }

        if (!string.IsNullOrWhiteSpace(Remediation)) {
            details["remediation"] = Remediation;
        }

        if (!string.IsNullOrWhiteSpace(Location)) {
            details["location"] = Location;
        }

        if (RunIndex.HasValue) {
            details["runIndex"] = RunIndex.Value.ToString(CultureInfo.InvariantCulture);
        }

        if (PageNumber.HasValue) {
            details["pageNumber"] = PageNumber.Value.ToString(CultureInfo.InvariantCulture);
        }

        if (TableRowIndex.HasValue && TableColumnIndex.HasValue) {
            details["tableRowIndex"] = TableRowIndex.Value.ToString(CultureInfo.InvariantCulture);
            details["tableColumnIndex"] = TableColumnIndex.Value.ToString(CultureInfo.InvariantCulture);
        }

        if (!string.IsNullOrWhiteSpace(FieldName)) {
            details["fieldName"] = FieldName;
        }

        var layoutDiagnostic = new PdfLayoutDiagnostic(
            IsControlCharacter ? PdfLayoutDiagnosticKind.SkippedContent : PdfLayoutDiagnosticKind.SimplifiedContent,
            Source,
            Message);

        return new PdfConversionWarning(
            converter,
            Code,
            Source,
            Message,
            PdfConversionWarningSeverity.Error,
            layoutDiagnostic,
            new ReadOnlyDictionary<string, string>(details));
    }

    /// <summary>
    /// Returns a copy of this diagnostic annotated with a one-based generated page number.
    /// </summary>
    /// <param name="pageNumber">One-based generated page number.</param>
    /// <returns>A diagnostic carrying the same text details and the supplied page number.</returns>
    public PdfTextEncodingDiagnostic WithPageNumber(int pageNumber) {
        Guard.PositiveInteger(pageNumber, nameof(pageNumber));
        return new PdfTextEncodingDiagnostic(Source, Index, CodePoint, Text, IsControlCharacter, Encoding, Remediation, Location, RunIndex, pageNumber, TableRowIndex, TableColumnIndex, FieldName);
    }

    /// <summary>
    /// Returns a copy of this diagnostic annotated with zero-based table cell coordinates.
    /// </summary>
    /// <param name="rowIndex">Zero-based table row index.</param>
    /// <param name="columnIndex">Zero-based table column index.</param>
    /// <returns>A diagnostic carrying the same text details and the supplied table cell coordinates.</returns>
    public PdfTextEncodingDiagnostic WithTableCell(int rowIndex, int columnIndex) {
        if (rowIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(rowIndex), rowIndex, "PDF diagnostic table row index must be zero or greater.");
        }

        if (columnIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, "PDF diagnostic table column index must be zero or greater.");
        }

        return new PdfTextEncodingDiagnostic(Source, Index, CodePoint, Text, IsControlCharacter, Encoding, Remediation, Location, RunIndex, PageNumber, rowIndex, columnIndex, FieldName);
    }

    /// <summary>
    /// Returns a copy of this diagnostic annotated with a generated AcroForm field name.
    /// </summary>
    /// <param name="fieldName">Generated field name.</param>
    /// <returns>A diagnostic carrying the same text details and the supplied field name.</returns>
    public PdfTextEncodingDiagnostic WithFieldName(string fieldName) {
        if (string.IsNullOrWhiteSpace(fieldName)) {
            return this;
        }

        return new PdfTextEncodingDiagnostic(Source, Index, CodePoint, Text, IsControlCharacter, Encoding, Remediation, Location, RunIndex, PageNumber, TableRowIndex, TableColumnIndex, fieldName);
    }

    private static string CreateMessage(int index, string codePoint, string text, bool isControlCharacter, string encoding, string remediation) {
        string indexText = index.ToString(CultureInfo.InvariantCulture);
        if (isControlCharacter) {
            return "Text contains control character " + codePoint + " at index " + indexText + ". PDF text output cannot render control characters directly; use paragraphs, line breaks, tables, or spacing primitives for layout.";
        }

        string rendered = string.IsNullOrEmpty(text) ? string.Empty : " '" + text + "'";
        if (!string.IsNullOrWhiteSpace(encoding)) {
            string suffix = string.IsNullOrWhiteSpace(remediation) ? string.Empty : " " + remediation;
            return "Text contains character " + codePoint + rendered + " at index " + indexText + " that cannot be encoded with " + encoding + "." + suffix;
        }

        return "Text contains character " + codePoint + rendered + " at index " + indexText + " that cannot be encoded with PDF WinAnsiEncoding. Embedded Unicode fonts are required for this text.";
    }
}
