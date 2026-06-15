namespace OfficeIMO.Rtf;

/// <summary>
/// Footnote, endnote, or annotation content in an RTF document.
/// </summary>
public sealed class RtfNote {
    private readonly List<RtfParagraph> _paragraphs = new List<RtfParagraph>();

    /// <summary>Creates a note with the specified kind.</summary>
    public RtfNote(RtfNoteKind kind) {
        Kind = kind;
    }

    /// <summary>Note destination kind.</summary>
    public RtfNoteKind Kind { get; }

    /// <summary>Optional note identifier from annotation metadata.</summary>
    public string? Id { get; set; }

    /// <summary>Optional annotation author.</summary>
    public string? Author { get; set; }

    /// <summary>Optional annotation creation time.</summary>
    public DateTime? Created { get; set; }

    /// <summary>Paragraphs contained by this note.</summary>
    public IReadOnlyList<RtfParagraph> Paragraphs => _paragraphs.AsReadOnly();

    /// <summary>Adds a paragraph to this note.</summary>
    public RtfParagraph AddParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _paragraphs.Add(paragraph);
        return paragraph;
    }

    /// <summary>Returns note text with paragraphs separated by new lines.</summary>
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
