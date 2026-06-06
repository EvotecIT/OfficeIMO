using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes text that cannot be written through the current generated PDF text encoding path.
/// </summary>
public sealed class PdfTextEncodingDiagnostic {
    /// <summary>
    /// Creates a text encoding diagnostic.
    /// </summary>
    /// <param name="source">Caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="index">UTF-16 index of the unsupported character in the inspected text.</param>
    /// <param name="codePoint">Unicode code point formatted as U+XXXX or U+XXXXX.</param>
    /// <param name="text">Unsupported character or surrogate pair text. Control characters are represented as an empty string.</param>
    /// <param name="isControlCharacter">Whether the unsupported character is a control character.</param>
    public PdfTextEncodingDiagnostic(string source, int index, string codePoint, string text, bool isControlCharacter) {
        Source = source ?? string.Empty;
        Index = index;
        CodePoint = codePoint ?? string.Empty;
        Text = text ?? string.Empty;
        IsControlCharacter = isControlCharacter;
        Message = CreateMessage(Index, CodePoint, Text, IsControlCharacter);
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

    /// <summary>Stable warning code suitable for shared conversion reports.</summary>
    public string Code => IsControlCharacter ? "unsupported-control-character" : "unsupported-text-glyph";

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

    private static string CreateMessage(int index, string codePoint, string text, bool isControlCharacter) {
        string indexText = index.ToString(CultureInfo.InvariantCulture);
        if (isControlCharacter) {
            return "Text contains control character " + codePoint + " at index " + indexText + ". PDF text output cannot render control characters directly; use paragraphs, line breaks, tables, or spacing primitives for layout.";
        }

        string rendered = string.IsNullOrEmpty(text) ? string.Empty : " '" + text + "'";
        return "Text contains character " + codePoint + rendered + " at index " + indexText + " that cannot be encoded with PDF WinAnsiEncoding. Embedded Unicode fonts are required for this text.";
    }
}
