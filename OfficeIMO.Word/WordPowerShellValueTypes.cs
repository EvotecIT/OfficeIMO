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

internal static class WordPowerShellValueTypeExtensions {
    internal static BreakValues ToOpenXml(this WordBreakType value) => value switch {
        WordBreakType.Page => BreakValues.Page,
        WordBreakType.Column => BreakValues.Column,
        WordBreakType.TextWrapping => BreakValues.TextWrapping,
        _ => throw Unsupported(value)
    };

    internal static HeaderFooterValues ToOpenXml(this WordHeaderFooterType value) => value switch {
        WordHeaderFooterType.Default => HeaderFooterValues.Default,
        WordHeaderFooterType.First => HeaderFooterValues.First,
        WordHeaderFooterType.Even => HeaderFooterValues.Even,
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

    internal static TableVerticalAlignmentValues ToOpenXml(this WordTableVerticalAlignment value) => value switch {
        WordTableVerticalAlignment.Top => TableVerticalAlignmentValues.Top,
        WordTableVerticalAlignment.Center => TableVerticalAlignmentValues.Center,
        WordTableVerticalAlignment.Bottom => TableVerticalAlignmentValues.Bottom,
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

    private static ArgumentOutOfRangeException Unsupported<T>(T value) where T : struct, Enum =>
        new(nameof(value), value, $"Unsupported {typeof(T).Name} value.");
}
