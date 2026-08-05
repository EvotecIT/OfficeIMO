using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Specifies the kind of break inserted into a Word document or paragraph.</summary>
public enum WordBreakType {
    /// <summary>Starts content on the next page.</summary>
    Page,
    /// <summary>Starts content in the next column.</summary>
    Column,
    /// <summary>Starts a new line without starting a new paragraph.</summary>
    TextWrapping
}

/// <summary>Specifies which header or footer variant is used.</summary>
public enum WordHeaderFooterType {
    /// <summary>The default header or footer.</summary>
    Default,
    /// <summary>The first-page header or footer.</summary>
    First,
    /// <summary>The even-page header or footer.</summary>
    Even
}

/// <summary>Specifies paragraph justification.</summary>
public enum WordParagraphAlignment {
    /// <summary>Aligns content to the left edge.</summary>
    Left,
    /// <summary>Aligns content to the logical start edge.</summary>
    Start,
    /// <summary>Centers content.</summary>
    Center,
    /// <summary>Aligns content to the right edge.</summary>
    Right,
    /// <summary>Aligns content to the logical end edge.</summary>
    End,
    /// <summary>Justifies content on both edges.</summary>
    Both,
    /// <summary>Uses medium Kashida justification.</summary>
    MediumKashida,
    /// <summary>Distributes content across the line.</summary>
    Distribute,
    /// <summary>Aligns content to a numbering tab.</summary>
    NumTab,
    /// <summary>Uses high Kashida justification.</summary>
    HighKashida,
    /// <summary>Uses low Kashida justification.</summary>
    LowKashida,
    /// <summary>Uses Thai distributed justification.</summary>
    ThaiDistribute
}

/// <summary>Specifies an underline style.</summary>
public enum WordUnderlineStyle {
    /// <summary>Single underline.</summary>
    Single,
    /// <summary>Underlines words only.</summary>
    Words,
    /// <summary>Double underline.</summary>
    Double,
    /// <summary>Thick underline.</summary>
    Thick,
    /// <summary>Dotted underline.</summary>
    Dotted,
    /// <summary>Heavy dotted underline.</summary>
    DottedHeavy,
    /// <summary>Dashed underline.</summary>
    Dash,
    /// <summary>Heavy dashed underline.</summary>
    DashedHeavy,
    /// <summary>Long-dash underline.</summary>
    DashLong,
    /// <summary>Heavy long-dash underline.</summary>
    DashLongHeavy,
    /// <summary>Dot-dash underline.</summary>
    DotDash,
    /// <summary>Heavy dot-dash underline.</summary>
    DashDotHeavy,
    /// <summary>Double-dot-dash underline.</summary>
    DotDotDash,
    /// <summary>Heavy double-dot-dash underline.</summary>
    DashDotDotHeavy,
    /// <summary>Wavy underline.</summary>
    Wave,
    /// <summary>Heavy wavy underline.</summary>
    WavyHeavy,
    /// <summary>Double wavy underline.</summary>
    WavyDouble,
    /// <summary>No underline.</summary>
    None
}

/// <summary>Specifies vertical alignment within a table cell.</summary>
public enum WordTableVerticalAlignment {
    /// <summary>Aligns content to the top.</summary>
    Top,
    /// <summary>Centers content vertically.</summary>
    Center,
    /// <summary>Aligns content to the bottom.</summary>
    Bottom
}

/// <summary>Specifies the editing restriction applied to a Word document.</summary>
public enum WordDocumentProtectionType {
    /// <summary>Does not restrict editing.</summary>
    None,
    /// <summary>Restricts the document to read-only use.</summary>
    ReadOnly,
    /// <summary>Allows comments only.</summary>
    Comments,
    /// <summary>Allows tracked changes only.</summary>
    TrackedChanges,
    /// <summary>Allows form-field editing only.</summary>
    Forms
}

/// <summary>Specifies the writing direction used by Word content.</summary>
public enum WordTextDirection {
    /// <summary>Left-to-right text arranged from top to bottom.</summary>
    LeftToRightTopToBottom,
    /// <summary>Office 2010 left-to-right top-to-bottom direction.</summary>
    LeftToRightTopToBottom2010,
    /// <summary>Top-to-bottom text arranged from right to left.</summary>
    TopToBottomRightToLeft,
    /// <summary>Office 2010 top-to-bottom right-to-left direction.</summary>
    TopToBottomRightToLeft2010,
    /// <summary>Bottom-to-top text arranged from left to right.</summary>
    BottomToTopLeftToRight,
    /// <summary>Office 2010 bottom-to-top left-to-right direction.</summary>
    BottomToTopLeftToRight2010,
    /// <summary>Rotated left-to-right top-to-bottom direction.</summary>
    LeftToRightTopToBottomRotated,
    /// <summary>Office 2010 rotated left-to-right top-to-bottom direction.</summary>
    LeftToRightTopToBottomRotated2010,
    /// <summary>Rotated top-to-bottom right-to-left direction.</summary>
    TopToBottomRightToLeftRotated,
    /// <summary>Office 2010 rotated top-to-bottom right-to-left direction.</summary>
    TopToBottomRightToLeftRotated2010,
    /// <summary>Rotated top-to-bottom left-to-right direction.</summary>
    TopToBottomLeftToRightRotated,
    /// <summary>Office 2010 rotated top-to-bottom left-to-right direction.</summary>
    TopToBottomLeftToRightRotated2010
}

/// <summary>Specifies the horizontal anchor used for positioned Word content.</summary>
public enum WordHorizontalRelativePosition {
    /// <summary>Anchors relative to the margins.</summary>
    Margin,
    /// <summary>Anchors relative to the page.</summary>
    Page,
    /// <summary>Anchors relative to the column.</summary>
    Column,
    /// <summary>Anchors relative to a character.</summary>
    Character,
    /// <summary>Anchors relative to the left margin.</summary>
    LeftMargin,
    /// <summary>Anchors relative to the right margin.</summary>
    RightMargin,
    /// <summary>Anchors relative to the inside margin.</summary>
    InsideMargin,
    /// <summary>Anchors relative to the outside margin.</summary>
    OutsideMargin
}

/// <summary>Specifies the vertical anchor used for positioned Word content.</summary>
public enum WordVerticalRelativePosition {
    /// <summary>Anchors relative to the margins.</summary>
    Margin,
    /// <summary>Anchors relative to the page.</summary>
    Page,
    /// <summary>Anchors relative to the paragraph.</summary>
    Paragraph,
    /// <summary>Anchors relative to the line.</summary>
    Line,
    /// <summary>Anchors relative to the top margin.</summary>
    TopMargin,
    /// <summary>Anchors relative to the bottom margin.</summary>
    BottomMargin,
    /// <summary>Anchors relative to the inside margin.</summary>
    InsideMargin,
    /// <summary>Anchors relative to the outside margin.</summary>
    OutsideMargin
}

internal static class WordValueTypeExtensions {
    internal static BreakValues ToOpenXml(this WordBreakType value) => value switch {
        WordBreakType.Page => BreakValues.Page,
        WordBreakType.Column => BreakValues.Column,
        WordBreakType.TextWrapping => BreakValues.TextWrapping,
        _ => throw Unsupported(value)
    };

    internal static WordBreakType ToOfficeEnum(this BreakValues value) => value switch {
        _ when value == BreakValues.Page => WordBreakType.Page,
        _ when value == BreakValues.Column => WordBreakType.Column,
        _ when value == BreakValues.TextWrapping => WordBreakType.TextWrapping,
        _ => throw Unsupported(value)
    };

    internal static HeaderFooterValues ToOpenXml(this WordHeaderFooterType value) => value switch {
        WordHeaderFooterType.Default => HeaderFooterValues.Default,
        WordHeaderFooterType.First => HeaderFooterValues.First,
        WordHeaderFooterType.Even => HeaderFooterValues.Even,
        _ => throw Unsupported(value)
    };

    internal static WordHeaderFooterType ToOfficeEnum(this HeaderFooterValues value) => value switch {
        _ when value == HeaderFooterValues.Default => WordHeaderFooterType.Default,
        _ when value == HeaderFooterValues.First => WordHeaderFooterType.First,
        _ when value == HeaderFooterValues.Even => WordHeaderFooterType.Even,
        _ => throw Unsupported(value)
    };

    internal static JustificationValues ToOpenXml(this WordParagraphAlignment value) => value switch {
        WordParagraphAlignment.Left => JustificationValues.Left,
        WordParagraphAlignment.Start => JustificationValues.Start,
        WordParagraphAlignment.Center => JustificationValues.Center,
        WordParagraphAlignment.Right => JustificationValues.Right,
        WordParagraphAlignment.End => JustificationValues.End,
        WordParagraphAlignment.Both => JustificationValues.Both,
        WordParagraphAlignment.MediumKashida => JustificationValues.MediumKashida,
        WordParagraphAlignment.Distribute => JustificationValues.Distribute,
        WordParagraphAlignment.NumTab => JustificationValues.NumTab,
        WordParagraphAlignment.HighKashida => JustificationValues.HighKashida,
        WordParagraphAlignment.LowKashida => JustificationValues.LowKashida,
        WordParagraphAlignment.ThaiDistribute => JustificationValues.ThaiDistribute,
        _ => throw Unsupported(value)
    };

    internal static WordParagraphAlignment ToOfficeEnum(this JustificationValues value) => value switch {
        _ when value == JustificationValues.Left => WordParagraphAlignment.Left,
        _ when value == JustificationValues.Start => WordParagraphAlignment.Start,
        _ when value == JustificationValues.Center => WordParagraphAlignment.Center,
        _ when value == JustificationValues.Right => WordParagraphAlignment.Right,
        _ when value == JustificationValues.End => WordParagraphAlignment.End,
        _ when value == JustificationValues.Both => WordParagraphAlignment.Both,
        _ when value == JustificationValues.MediumKashida => WordParagraphAlignment.MediumKashida,
        _ when value == JustificationValues.Distribute => WordParagraphAlignment.Distribute,
        _ when value == JustificationValues.NumTab => WordParagraphAlignment.NumTab,
        _ when value == JustificationValues.HighKashida => WordParagraphAlignment.HighKashida,
        _ when value == JustificationValues.LowKashida => WordParagraphAlignment.LowKashida,
        _ when value == JustificationValues.ThaiDistribute => WordParagraphAlignment.ThaiDistribute,
        _ => throw Unsupported(value)
    };

    internal static UnderlineValues ToOpenXml(this WordUnderlineStyle value) => value switch {
        WordUnderlineStyle.Single => UnderlineValues.Single,
        WordUnderlineStyle.Words => UnderlineValues.Words,
        WordUnderlineStyle.Double => UnderlineValues.Double,
        WordUnderlineStyle.Thick => UnderlineValues.Thick,
        WordUnderlineStyle.Dotted => UnderlineValues.Dotted,
        WordUnderlineStyle.DottedHeavy => UnderlineValues.DottedHeavy,
        WordUnderlineStyle.Dash => UnderlineValues.Dash,
        WordUnderlineStyle.DashedHeavy => UnderlineValues.DashedHeavy,
        WordUnderlineStyle.DashLong => UnderlineValues.DashLong,
        WordUnderlineStyle.DashLongHeavy => UnderlineValues.DashLongHeavy,
        WordUnderlineStyle.DotDash => UnderlineValues.DotDash,
        WordUnderlineStyle.DashDotHeavy => UnderlineValues.DashDotHeavy,
        WordUnderlineStyle.DotDotDash => UnderlineValues.DotDotDash,
        WordUnderlineStyle.DashDotDotHeavy => UnderlineValues.DashDotDotHeavy,
        WordUnderlineStyle.Wave => UnderlineValues.Wave,
        WordUnderlineStyle.WavyHeavy => UnderlineValues.WavyHeavy,
        WordUnderlineStyle.WavyDouble => UnderlineValues.WavyDouble,
        WordUnderlineStyle.None => UnderlineValues.None,
        _ => throw Unsupported(value)
    };

    internal static WordUnderlineStyle ToOfficeEnum(this UnderlineValues value) => value switch {
        _ when value == UnderlineValues.Single => WordUnderlineStyle.Single,
        _ when value == UnderlineValues.Words => WordUnderlineStyle.Words,
        _ when value == UnderlineValues.Double => WordUnderlineStyle.Double,
        _ when value == UnderlineValues.Thick => WordUnderlineStyle.Thick,
        _ when value == UnderlineValues.Dotted => WordUnderlineStyle.Dotted,
        _ when value == UnderlineValues.DottedHeavy => WordUnderlineStyle.DottedHeavy,
        _ when value == UnderlineValues.Dash => WordUnderlineStyle.Dash,
        _ when value == UnderlineValues.DashedHeavy => WordUnderlineStyle.DashedHeavy,
        _ when value == UnderlineValues.DashLong => WordUnderlineStyle.DashLong,
        _ when value == UnderlineValues.DashLongHeavy => WordUnderlineStyle.DashLongHeavy,
        _ when value == UnderlineValues.DotDash => WordUnderlineStyle.DotDash,
        _ when value == UnderlineValues.DashDotHeavy => WordUnderlineStyle.DashDotHeavy,
        _ when value == UnderlineValues.DotDotDash => WordUnderlineStyle.DotDotDash,
        _ when value == UnderlineValues.DashDotDotHeavy => WordUnderlineStyle.DashDotDotHeavy,
        _ when value == UnderlineValues.Wave => WordUnderlineStyle.Wave,
        _ when value == UnderlineValues.WavyHeavy => WordUnderlineStyle.WavyHeavy,
        _ when value == UnderlineValues.WavyDouble => WordUnderlineStyle.WavyDouble,
        _ when value == UnderlineValues.None => WordUnderlineStyle.None,
        _ => throw Unsupported(value)
    };

    internal static TableVerticalAlignmentValues ToOpenXml(this WordTableVerticalAlignment value) => value switch {
        WordTableVerticalAlignment.Top => TableVerticalAlignmentValues.Top,
        WordTableVerticalAlignment.Center => TableVerticalAlignmentValues.Center,
        WordTableVerticalAlignment.Bottom => TableVerticalAlignmentValues.Bottom,
        _ => throw Unsupported(value)
    };

    internal static WordTableVerticalAlignment ToOfficeEnum(this TableVerticalAlignmentValues value) => value switch {
        _ when value == TableVerticalAlignmentValues.Top => WordTableVerticalAlignment.Top,
        _ when value == TableVerticalAlignmentValues.Center => WordTableVerticalAlignment.Center,
        _ when value == TableVerticalAlignmentValues.Bottom => WordTableVerticalAlignment.Bottom,
        _ => throw Unsupported(value)
    };

    internal static DocumentProtectionValues ToOpenXml(this WordDocumentProtectionType value) => value switch {
        WordDocumentProtectionType.None => DocumentProtectionValues.None,
        WordDocumentProtectionType.ReadOnly => DocumentProtectionValues.ReadOnly,
        WordDocumentProtectionType.Comments => DocumentProtectionValues.Comments,
        WordDocumentProtectionType.TrackedChanges => DocumentProtectionValues.TrackedChanges,
        WordDocumentProtectionType.Forms => DocumentProtectionValues.Forms,
        _ => throw Unsupported(value)
    };

    internal static WordDocumentProtectionType ToOfficeEnum(this DocumentProtectionValues value) => value switch {
        _ when value == DocumentProtectionValues.None => WordDocumentProtectionType.None,
        _ when value == DocumentProtectionValues.ReadOnly => WordDocumentProtectionType.ReadOnly,
        _ when value == DocumentProtectionValues.Comments => WordDocumentProtectionType.Comments,
        _ when value == DocumentProtectionValues.TrackedChanges => WordDocumentProtectionType.TrackedChanges,
        _ when value == DocumentProtectionValues.Forms => WordDocumentProtectionType.Forms,
        _ => throw Unsupported(value)
    };

    internal static TextDirectionValues ToOpenXml(this WordTextDirection value) => value switch {
        WordTextDirection.LeftToRightTopToBottom => TextDirectionValues.LefToRightTopToBottom,
        WordTextDirection.LeftToRightTopToBottom2010 => TextDirectionValues.LeftToRightTopToBottom2010,
        WordTextDirection.TopToBottomRightToLeft => TextDirectionValues.TopToBottomRightToLeft,
        WordTextDirection.TopToBottomRightToLeft2010 => TextDirectionValues.TopToBottomRightToLeft2010,
        WordTextDirection.BottomToTopLeftToRight => TextDirectionValues.BottomToTopLeftToRight,
        WordTextDirection.BottomToTopLeftToRight2010 => TextDirectionValues.BottomToTopLeftToRight2010,
        WordTextDirection.LeftToRightTopToBottomRotated => TextDirectionValues.LefttoRightTopToBottomRotated,
        WordTextDirection.LeftToRightTopToBottomRotated2010 => TextDirectionValues.LeftToRightTopToBottomRotated2010,
        WordTextDirection.TopToBottomRightToLeftRotated => TextDirectionValues.TopToBottomRightToLeftRotated,
        WordTextDirection.TopToBottomRightToLeftRotated2010 => TextDirectionValues.TopToBottomRightToLeftRotated2010,
        WordTextDirection.TopToBottomLeftToRightRotated => TextDirectionValues.TopToBottomLeftToRightRotated,
        WordTextDirection.TopToBottomLeftToRightRotated2010 => TextDirectionValues.TopToBottomLeftToRightRotated2010,
        _ => throw Unsupported(value)
    };

    internal static WordTextDirection ToOfficeEnum(this TextDirectionValues value) => value switch {
        _ when value == TextDirectionValues.LefToRightTopToBottom => WordTextDirection.LeftToRightTopToBottom,
        _ when value == TextDirectionValues.LeftToRightTopToBottom2010 => WordTextDirection.LeftToRightTopToBottom2010,
        _ when value == TextDirectionValues.TopToBottomRightToLeft => WordTextDirection.TopToBottomRightToLeft,
        _ when value == TextDirectionValues.TopToBottomRightToLeft2010 => WordTextDirection.TopToBottomRightToLeft2010,
        _ when value == TextDirectionValues.BottomToTopLeftToRight => WordTextDirection.BottomToTopLeftToRight,
        _ when value == TextDirectionValues.BottomToTopLeftToRight2010 => WordTextDirection.BottomToTopLeftToRight2010,
        _ when value == TextDirectionValues.LefttoRightTopToBottomRotated => WordTextDirection.LeftToRightTopToBottomRotated,
        _ when value == TextDirectionValues.LeftToRightTopToBottomRotated2010 => WordTextDirection.LeftToRightTopToBottomRotated2010,
        _ when value == TextDirectionValues.TopToBottomRightToLeftRotated => WordTextDirection.TopToBottomRightToLeftRotated,
        _ when value == TextDirectionValues.TopToBottomRightToLeftRotated2010 => WordTextDirection.TopToBottomRightToLeftRotated2010,
        _ when value == TextDirectionValues.TopToBottomLeftToRightRotated => WordTextDirection.TopToBottomLeftToRightRotated,
        _ when value == TextDirectionValues.TopToBottomLeftToRightRotated2010 => WordTextDirection.TopToBottomLeftToRightRotated2010,
        _ => throw Unsupported(value)
    };

    internal static HorizontalRelativePositionValues ToOpenXml(this WordHorizontalRelativePosition value) => value switch {
        WordHorizontalRelativePosition.Margin => HorizontalRelativePositionValues.Margin,
        WordHorizontalRelativePosition.Page => HorizontalRelativePositionValues.Page,
        WordHorizontalRelativePosition.Column => HorizontalRelativePositionValues.Column,
        WordHorizontalRelativePosition.Character => HorizontalRelativePositionValues.Character,
        WordHorizontalRelativePosition.LeftMargin => HorizontalRelativePositionValues.LeftMargin,
        WordHorizontalRelativePosition.RightMargin => HorizontalRelativePositionValues.RightMargin,
        WordHorizontalRelativePosition.InsideMargin => HorizontalRelativePositionValues.InsideMargin,
        WordHorizontalRelativePosition.OutsideMargin => HorizontalRelativePositionValues.OutsideMargin,
        _ => throw Unsupported(value)
    };

    internal static WordHorizontalRelativePosition ToOfficeEnum(this HorizontalRelativePositionValues value) => value switch {
        _ when value == HorizontalRelativePositionValues.Margin => WordHorizontalRelativePosition.Margin,
        _ when value == HorizontalRelativePositionValues.Page => WordHorizontalRelativePosition.Page,
        _ when value == HorizontalRelativePositionValues.Column => WordHorizontalRelativePosition.Column,
        _ when value == HorizontalRelativePositionValues.Character => WordHorizontalRelativePosition.Character,
        _ when value == HorizontalRelativePositionValues.LeftMargin => WordHorizontalRelativePosition.LeftMargin,
        _ when value == HorizontalRelativePositionValues.RightMargin => WordHorizontalRelativePosition.RightMargin,
        _ when value == HorizontalRelativePositionValues.InsideMargin => WordHorizontalRelativePosition.InsideMargin,
        _ when value == HorizontalRelativePositionValues.OutsideMargin => WordHorizontalRelativePosition.OutsideMargin,
        _ => throw Unsupported(value)
    };

    internal static VerticalRelativePositionValues ToOpenXml(this WordVerticalRelativePosition value) => value switch {
        WordVerticalRelativePosition.Margin => VerticalRelativePositionValues.Margin,
        WordVerticalRelativePosition.Page => VerticalRelativePositionValues.Page,
        WordVerticalRelativePosition.Paragraph => VerticalRelativePositionValues.Paragraph,
        WordVerticalRelativePosition.Line => VerticalRelativePositionValues.Line,
        WordVerticalRelativePosition.TopMargin => VerticalRelativePositionValues.TopMargin,
        WordVerticalRelativePosition.BottomMargin => VerticalRelativePositionValues.BottomMargin,
        WordVerticalRelativePosition.InsideMargin => VerticalRelativePositionValues.InsideMargin,
        WordVerticalRelativePosition.OutsideMargin => VerticalRelativePositionValues.OutsideMargin,
        _ => throw Unsupported(value)
    };

    internal static WordVerticalRelativePosition ToOfficeEnum(this VerticalRelativePositionValues value) => value switch {
        _ when value == VerticalRelativePositionValues.Margin => WordVerticalRelativePosition.Margin,
        _ when value == VerticalRelativePositionValues.Page => WordVerticalRelativePosition.Page,
        _ when value == VerticalRelativePositionValues.Paragraph => WordVerticalRelativePosition.Paragraph,
        _ when value == VerticalRelativePositionValues.Line => WordVerticalRelativePosition.Line,
        _ when value == VerticalRelativePositionValues.TopMargin => WordVerticalRelativePosition.TopMargin,
        _ when value == VerticalRelativePositionValues.BottomMargin => WordVerticalRelativePosition.BottomMargin,
        _ when value == VerticalRelativePositionValues.InsideMargin => WordVerticalRelativePosition.InsideMargin,
        _ when value == VerticalRelativePositionValues.OutsideMargin => WordVerticalRelativePosition.OutsideMargin,
        _ => throw Unsupported(value)
    };

    internal static BreakValues? ToOpenXml(this WordBreakType? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static HeaderFooterValues? ToOpenXml(this WordHeaderFooterType? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static JustificationValues? ToOpenXml(this WordParagraphAlignment? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static UnderlineValues? ToOpenXml(this WordUnderlineStyle? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static TableVerticalAlignmentValues? ToOpenXml(this WordTableVerticalAlignment? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static DocumentProtectionValues? ToOpenXml(this WordDocumentProtectionType? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static TextDirectionValues? ToOpenXml(this WordTextDirection? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static HorizontalRelativePositionValues? ToOpenXml(this WordHorizontalRelativePosition? value) => value.HasValue ? value.Value.ToOpenXml() : null;
    internal static VerticalRelativePositionValues? ToOpenXml(this WordVerticalRelativePosition? value) => value.HasValue ? value.Value.ToOpenXml() : null;

    internal static WordBreakType? ToOfficeEnum(this BreakValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordHeaderFooterType? ToOfficeEnum(this HeaderFooterValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordParagraphAlignment? ToOfficeEnum(this JustificationValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordUnderlineStyle? ToOfficeEnum(this UnderlineValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordTableVerticalAlignment? ToOfficeEnum(this TableVerticalAlignmentValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordDocumentProtectionType? ToOfficeEnum(this DocumentProtectionValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordTextDirection? ToOfficeEnum(this TextDirectionValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordHorizontalRelativePosition? ToOfficeEnum(this HorizontalRelativePositionValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;
    internal static WordVerticalRelativePosition? ToOfficeEnum(this VerticalRelativePositionValues? value) => value.HasValue ? value.Value.ToOfficeEnum() : null;

    private static ArgumentOutOfRangeException Unsupported<T>(T value) where T : struct =>
        new(nameof(value), value, $"Unsupported {typeof(T).Name} value.");
}
