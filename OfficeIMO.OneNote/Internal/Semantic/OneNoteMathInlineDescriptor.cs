namespace OfficeIMO.OneNote;

internal sealed class OneNoteMathInlineDescriptor {
    internal uint Type { get; set; }
    internal uint? Count { get; set; }
    internal byte? Column { get; set; }
    internal byte? Alignment { get; set; }
    internal ushort? Character { get; set; }
    internal ushort? Character1 { get; set; }
    internal ushort? Character2 { get; set; }

    internal OneNoteMathInlineDescriptor Clone() => new OneNoteMathInlineDescriptor {
        Type = Type,
        Count = Count,
        Column = Column,
        Alignment = Alignment,
        Character = Character,
        Character1 = Character1,
        Character2 = Character2
    };
}

internal static class OneNoteMathRunPreservation {
    internal static OneNoteTextRun Clone(OneNoteTextRun source) {
        var clone = new OneNoteTextRun {
            Text = source.Text,
            MathExpression = source.MathExpression,
            MathDescriptor = source.MathDescriptor?.Clone(),
            Hyperlink = source.Hyperlink,
            HyperlinkProtected = source.HyperlinkProtected,
            StyleObjectId = source.StyleObjectId
        };
        CopyStyle(source.Style, clone.Style);
        foreach (OneNoteOpaqueProperty property in source.UnknownProperties) clone.UnknownProperties.Add(property);
        return clone;
    }

    internal static bool CanReuse(OneNoteTextRun semantic) {
        return semantic.MathExpression != null && semantic.PreservedMathExpression != null &&
            semantic.MathExpression.Equals(semantic.PreservedMathExpression) &&
            semantic.PreservedNativeMathRuns != null && semantic.PreservedNativeMathRuns.Count > 0;
    }

    internal static IReadOnlyList<OneNoteTextRun> CloneForWrite(OneNoteTextRun semantic) {
        if (!CanReuse(semantic)) throw new InvalidOperationException("The native math runs cannot be reused after the expression changes.");
        OneNoteTextRun first = semantic.PreservedNativeMathRuns![0];
        bool presentationChanged = !string.Equals(semantic.Hyperlink, first.Hyperlink, StringComparison.Ordinal) ||
            semantic.HyperlinkProtected != first.HyperlinkProtected || !StylesEqual(semantic.Style, first.Style);
        var clones = new List<OneNoteTextRun>(semantic.PreservedNativeMathRuns.Count);
        foreach (OneNoteTextRun source in semantic.PreservedNativeMathRuns) {
            OneNoteTextRun clone = Clone(source);
            if (presentationChanged) {
                clone.Hyperlink = semantic.Hyperlink;
                clone.HyperlinkProtected = semantic.HyperlinkProtected;
                clone.StyleObjectId = semantic.StyleObjectId;
                CopyStyle(semantic.Style, clone.Style);
                clone.Style.IsMath = true;
            }
            clones.Add(clone);
        }
        return clones;
    }

    private static void CopyStyle(OneNoteTextStyle source, OneNoteTextStyle destination) {
        destination.FontFamily = source.FontFamily;
        destination.FontSize = source.FontSize;
        destination.ColorArgb = source.ColorArgb;
        destination.HighlightColorArgb = source.HighlightColorArgb;
        destination.Bold = source.Bold;
        destination.Italic = source.Italic;
        destination.Underline = source.Underline;
        destination.Strikethrough = source.Strikethrough;
        destination.Superscript = source.Superscript;
        destination.Subscript = source.Subscript;
        destination.LanguageId = source.LanguageId;
        destination.IsMath = source.IsMath;
    }

    private static bool StylesEqual(OneNoteTextStyle left, OneNoteTextStyle right) =>
        string.Equals(left.FontFamily, right.FontFamily, StringComparison.Ordinal) &&
        left.FontSize == right.FontSize && left.ColorArgb == right.ColorArgb &&
        left.HighlightColorArgb == right.HighlightColorArgb && left.Bold == right.Bold &&
        left.Italic == right.Italic && left.Underline == right.Underline &&
        left.Strikethrough == right.Strikethrough && left.Superscript == right.Superscript &&
        left.Subscript == right.Subscript && left.LanguageId == right.LanguageId && left.IsMath == right.IsMath;
}
