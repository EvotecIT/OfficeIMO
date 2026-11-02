using System;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Keeps inline and list-item binding within the current slide or column and document.</summary>
internal sealed class PowerPointTextScope {
    internal PowerPointTextScope(PowerPointContentScope owner, PowerPointParagraph? paragraph = null, PowerPointTextBox? list = null, Action<PowerPointParagraph>? listItemDefaults = null) {
        Owner = owner; Document = owner.Document; TargetSlide = owner.NativeSlide; Paragraph = paragraph; List = list; ListItemDefaults = listItemDefaults;
    }
    internal PowerPointTextScope(PowerPointDocument document, PowerPointTextBox box, bool list = false, Action<PowerPointParagraph>? listItemDefaults = null) {
        Document = document; Box = list ? null : box; List = list ? box : null; ListItemDefaults = listItemDefaults; TargetSlide = box.OwnerSlide;
    }
    internal PowerPointTextScope(PowerPointDocument document, PowerPointSlide slide, PowerPointTableCell cell) {
        Document = document; Cell = cell; TargetSlide = slide;
    }
    internal PowerPointTextScope(PowerPointTextScope parent, PowerPointParagraph paragraph) {
        Owner = parent.Owner; Document = parent.Document; TargetSlide = parent.TargetSlide; Paragraph = paragraph;
    }
    internal PowerPointContentScope? Owner { get; }
    internal PowerPointDocument Document { get; }
    internal PowerPointSlide? TargetSlide { get; }
    internal OfficeTheme Theme => TargetSlide?.AuthoringTheme ?? Document.LayoutTheme.Theme;
    internal PowerPointParagraph? Paragraph { get; }
    internal PowerPointTextBox? List { get; }
    internal Action<PowerPointParagraph>? ListItemDefaults { get; }
    internal PowerPointTextBox? Box { get; }
    internal PowerPointTableCell? Cell { get; }
    internal bool HasTextBody => Box != null || Cell != null;
    private PowerPointParagraph? _bodyParagraph;
    private bool _bodyWritten;

    internal PowerPointParagraph StartParagraph() {
        if (!HasTextBody) throw new InvalidOperationException("This text scope does not own a text body.");
        var paragraphs = Box != null ? Box.Paragraphs : Cell!.Paragraphs;
        _bodyParagraph = !_bodyWritten && paragraphs.Count > 0 ? paragraphs[0] :
            Box != null ? Box.AddParagraph() : Cell!.AddParagraph();
        _bodyWritten = true;
        return _bodyParagraph;
    }

    internal PowerPointParagraph TextParagraph() => Paragraph ?? _bodyParagraph ?? StartParagraph();
    internal static PowerPointTextScope? Current {
        get {
            var scope = OfficeDocumentContext.Find<PowerPointTextScope>();
            return scope != null && ReferenceEquals(scope.Document, OfficeDocumentContext.Find<PowerPointDocument>()) &&
                (scope.Owner == null || ReferenceEquals(scope.Owner, PowerPointContentScope.Current)) ? scope : null;
        }
    }
}

/// <summary>Maps shared text vocabulary to the native PowerPoint run and theme.</summary>
internal static class PowerPointTextFormatting {
    internal static void Apply(PowerPointTextRun run, PowerPointContentScope scope, bool bold, bool italic,
        bool strikethrough, PowerPointUnderlineStyle? underline, OfficeColor? color, OfficeColor? highlight,
        string? font, OfficeLength? size, string? link = null) =>
        Apply(run, scope.Document, scope.NativeSlide.AuthoringTheme, bold, italic, strikethrough, underline, color, highlight, font, size, link);

    internal static void Apply(PowerPointTextRun run, PowerPointTextScope scope, bool bold, bool italic,
        bool strikethrough, PowerPointUnderlineStyle? underline, OfficeColor? color, OfficeColor? highlight,
        string? font, OfficeLength? size, string? link = null) =>
        Apply(run, scope.Document, scope.Theme, bold, italic, strikethrough, underline, color, highlight, font, size, link);

    private static void Apply(PowerPointTextRun run, PowerPointDocument document, OfficeTheme theme, bool bold, bool italic,
        bool strikethrough, PowerPointUnderlineStyle? underline, OfficeColor? color, OfficeColor? highlight,
        string? font, OfficeLength? size, string? link) {
        if (bold) run.Bold = true;
        if (italic) run.Italic = true;
        if (strikethrough) run.Strikethrough = true;
        if (underline.HasValue) run.UnderlineStyle = underline;
        var context = new PowerPointLayoutContext(document, PowerPointLayoutTheme.From(theme));
        if (color.HasValue) run.Color = context.ResolveColor(color);
        if (highlight.HasValue) run.HighlightColor = context.ResolveColor(highlight);
        if (font != null) run.FontName = font;
        if (size.HasValue) run.FontSize = size.Value;
        if (link != null) run.SetHyperlink(link);
    }
    internal static PowerPointTextAlignment Alignment(OfficeAlignment alignment) => alignment switch {
        OfficeAlignment.Center => PowerPointTextAlignment.Center,
        OfficeAlignment.Right or OfficeAlignment.End => PowerPointTextAlignment.Right,
        OfficeAlignment.Justify => PowerPointTextAlignment.Justified,
        _ => PowerPointTextAlignment.Left
    };
}
