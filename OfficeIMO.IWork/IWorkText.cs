namespace OfficeIMO.IWork;

/// <summary>Horizontal alignment recovered from an iWork paragraph style.</summary>
public enum IWorkTextAlignment {
    Natural,
    Left,
    Center,
    Right,
    Justified
}

/// <summary>The source delimiter that ended an iWork text paragraph.</summary>
public enum IWorkParagraphBreakKind {
    None,
    Paragraph,
    Section,
    Layout,
    Page
}

/// <summary>An immutable RGB color recovered from an iWork style.</summary>
public sealed class IWorkColor {
    internal IWorkColor(byte red, byte green, byte blue, byte alpha) {
        Red = red;
        Green = green;
        Blue = blue;
        Alpha = alpha;
    }

    /// <summary>Gets the red component.</summary>
    public byte Red { get; }
    /// <summary>Gets the green component.</summary>
    public byte Green { get; }
    /// <summary>Gets the blue component.</summary>
    public byte Blue { get; }
    /// <summary>Gets the alpha component.</summary>
    public byte Alpha { get; }
    /// <summary>Gets the opaque RGB representation used by Office formats.</summary>
    public string RgbHex => $"{Red:X2}{Green:X2}{Blue:X2}";
}

/// <summary>Character formatting recovered from an iWork text style.</summary>
public sealed class IWorkTextStyle {
    internal IWorkTextStyle(string? name, bool? bold, bool? italic, bool? underline,
        bool? strikethrough, double? fontSizePoints, string? fontName,
        IWorkColor? color, IWorkColor? backgroundColor) {
        Name = name;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Strikethrough = strikethrough;
        FontSizePoints = fontSizePoints;
        FontName = fontName;
        Color = color;
        BackgroundColor = backgroundColor;
    }

    /// <summary>Gets the source style name, when present.</summary>
    public string? Name { get; }
    /// <summary>Gets the explicit bold setting.</summary>
    public bool? Bold { get; }
    /// <summary>Gets the explicit italic setting.</summary>
    public bool? Italic { get; }
    /// <summary>Gets the explicit underline setting.</summary>
    public bool? Underline { get; }
    /// <summary>Gets the explicit strikethrough setting.</summary>
    public bool? Strikethrough { get; }
    /// <summary>Gets the font size in points.</summary>
    public double? FontSizePoints { get; }
    /// <summary>Gets the font family name.</summary>
    public string? FontName { get; }
    /// <summary>Gets the foreground color.</summary>
    public IWorkColor? Color { get; }
    /// <summary>Gets the background or highlight color.</summary>
    public IWorkColor? BackgroundColor { get; }
}

/// <summary>Paragraph formatting recovered from an iWork paragraph style.</summary>
public sealed class IWorkParagraphStyle {
    internal IWorkParagraphStyle(string? name, IWorkTextAlignment? alignment,
        double? firstLineIndentPoints, double? leftIndentPoints, double? rightIndentPoints,
        double? spaceBeforePoints, double? spaceAfterPoints, bool? pageBreakBefore,
        bool? keepWithNext, bool? keepLinesTogether, IWorkTextStyle textStyle) {
        Name = name;
        Alignment = alignment;
        FirstLineIndentPoints = firstLineIndentPoints;
        LeftIndentPoints = leftIndentPoints;
        RightIndentPoints = rightIndentPoints;
        SpaceBeforePoints = spaceBeforePoints;
        SpaceAfterPoints = spaceAfterPoints;
        PageBreakBefore = pageBreakBefore;
        KeepWithNext = keepWithNext;
        KeepLinesTogether = keepLinesTogether;
        TextStyle = textStyle;
    }

    /// <summary>Gets the source style name, when present.</summary>
    public string? Name { get; }
    /// <summary>Gets paragraph alignment.</summary>
    public IWorkTextAlignment? Alignment { get; }
    /// <summary>Gets first-line indentation in points.</summary>
    public double? FirstLineIndentPoints { get; }
    /// <summary>Gets left indentation in points.</summary>
    public double? LeftIndentPoints { get; }
    /// <summary>Gets right indentation in points.</summary>
    public double? RightIndentPoints { get; }
    /// <summary>Gets spacing before the paragraph in points.</summary>
    public double? SpaceBeforePoints { get; }
    /// <summary>Gets spacing after the paragraph in points.</summary>
    public double? SpaceAfterPoints { get; }
    /// <summary>Gets whether the paragraph starts on a new page.</summary>
    public bool? PageBreakBefore { get; }
    /// <summary>Gets whether the paragraph should stay with the next paragraph.</summary>
    public bool? KeepWithNext { get; }
    /// <summary>Gets whether the paragraph lines should stay together.</summary>
    public bool? KeepLinesTogether { get; }
    /// <summary>Gets character defaults carried by the paragraph style.</summary>
    public IWorkTextStyle TextStyle { get; }
}

/// <summary>One contiguous rich-text run.</summary>
public sealed class IWorkTextRun {
    internal IWorkTextRun(string text, IWorkTextStyle style, string? hyperlink) {
        Text = text;
        Style = style;
        Hyperlink = hyperlink;
    }

    /// <summary>Gets run text.</summary>
    public string Text { get; }
    /// <summary>Gets resolved run formatting.</summary>
    public IWorkTextStyle Style { get; }
    /// <summary>Gets an external hyperlink target, when present.</summary>
    public string? Hyperlink { get; }
}

/// <summary>One paragraph in an iWork rich-text storage.</summary>
public sealed class IWorkTextParagraph {
    internal IWorkTextParagraph(IReadOnlyList<IWorkTextRun> runs, IWorkParagraphStyle style,
        int listLevel, string? listLabel, IWorkParagraphBreakKind breakKind) {
        Runs = Array.AsReadOnly(runs.ToArray());
        Style = style;
        ListLevel = listLevel;
        ListLabel = listLabel;
        BreakKind = breakKind;
    }

    /// <summary>Gets rich-text runs in source order.</summary>
    public IReadOnlyList<IWorkTextRun> Runs { get; }
    /// <summary>Gets resolved paragraph formatting.</summary>
    public IWorkParagraphStyle Style { get; }
    /// <summary>Gets the zero-based list level, or -1 for a non-list paragraph.</summary>
    public int ListLevel { get; }
    /// <summary>Gets the source list label, when directly recoverable.</summary>
    public string? ListLabel { get; }
    /// <summary>Gets the delimiter that ended the paragraph.</summary>
    public IWorkParagraphBreakKind BreakKind { get; }
    /// <summary>Gets paragraph text without its terminal delimiter.</summary>
    public string Text => string.Concat(Runs.Select(run => run.Text));
}

/// <summary>Immutable rich text recovered from one iWork text storage.</summary>
public sealed class IWorkTextContent {
    internal IWorkTextContent(IReadOnlyList<IWorkTextParagraph> paragraphs, bool isComplete) {
        Paragraphs = Array.AsReadOnly(paragraphs.ToArray());
        IsComplete = isComplete;
    }

    /// <summary>Gets paragraphs, including meaningful empty paragraphs.</summary>
    public IReadOnlyList<IWorkTextParagraph> Paragraphs { get; }
    /// <summary>Gets whether text and all referenced style records were decoded.</summary>
    public bool IsComplete { get; }
    /// <summary>Gets normalized plain text while preserving paragraph boundaries.</summary>
    public string PlainText => string.Join("\n", Paragraphs.Select(paragraph => paragraph.Text));
}

/// <summary>Headers and footers associated with one Pages section in source order.</summary>
public sealed class IWorkPagesSection {
    internal IWorkPagesSection(int index, IReadOnlyList<IWorkTextContent> headers,
        IReadOnlyList<IWorkTextContent> footers) {
        Index = index;
        HeaderContents = Array.AsReadOnly(headers.ToArray());
        FooterContents = Array.AsReadOnly(footers.ToArray());
    }

    /// <summary>Gets the zero-based source section index.</summary>
    public int Index { get; }
    /// <summary>Gets rich header storages associated with this section.</summary>
    public IReadOnlyList<IWorkTextContent> HeaderContents { get; }
    /// <summary>Gets rich footer storages associated with this section.</summary>
    public IReadOnlyList<IWorkTextContent> FooterContents { get; }
}
