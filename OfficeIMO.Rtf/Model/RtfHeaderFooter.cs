namespace OfficeIMO.Rtf;

/// <summary>
/// Header or footer content for an RTF document.
/// </summary>
public sealed class RtfHeaderFooter {
    private readonly List<RtfParagraph> _paragraphs = new List<RtfParagraph>();

    internal RtfHeaderFooter(RtfHeaderFooterKind kind) {
        Kind = kind;
    }

    /// <summary>Header or footer destination kind.</summary>
    public RtfHeaderFooterKind Kind { get; }

    /// <summary>Paragraphs contained by this header or footer.</summary>
    public IReadOnlyList<RtfParagraph> Paragraphs => _paragraphs.AsReadOnly();

    /// <summary>Adds a paragraph to this header or footer.</summary>
    public RtfParagraph AddParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _paragraphs.Add(paragraph);
        return paragraph;
    }

    /// <summary>Returns the header or footer text with paragraphs separated by new lines.</summary>
    public string ToPlainText() {
        var builder = new StringBuilder();
        for (int i = 0; i < _paragraphs.Count; i++) {
            if (i > 0) {
                builder.AppendLine();
            }

            builder.Append(_paragraphs[i].ToPlainText());
        }

        return builder.ToString();
    }

    internal void AddParsedParagraph(RtfParagraph paragraph) {
        _paragraphs.Add(paragraph ?? throw new ArgumentNullException(nameof(paragraph)));
    }
}
