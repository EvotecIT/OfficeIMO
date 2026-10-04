namespace OfficeIMO.AsciiDoc;

/// <summary>Keeps scalar assignments and typed inline edits on one effective content sequence.</summary>
internal sealed class AsciiDocEditableInlineContent {
    private readonly string _originalText;
    private bool _assigned;

    internal AsciiDocEditableInlineContent(string text, AsciiDocInlineSequence inlines) {
        _originalText = text;
        Inlines = inlines;
    }

    internal AsciiDocInlineSequence Inlines { get; private set; }
    internal string Text => _assigned || Inlines.IsModified ? Inlines.ToAsciiDoc() : _originalText;
    internal bool IsModified => _assigned || Inlines.IsModified;

    internal void Assign(string value) {
        if (string.Equals(Text, value, StringComparison.Ordinal)) return;
        AsciiDocInlineSequence parsed = AsciiDocInlineSequence.Parse(value);
        Inlines = parsed;
        _assigned = true;
    }

    internal string Write(AsciiDocWriterContext context) => Inlines.Write(context);
}
