using System;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Keeps inline and list-item binding within the current slide or column and document.</summary>
internal sealed class PowerPointTextScope {
    internal PowerPointTextScope(PowerPointContentScope owner, PowerPointParagraph? paragraph = null, PowerPointTextBox? list = null) {
        Owner = owner; Paragraph = paragraph; List = list;
    }
    internal PowerPointContentScope Owner { get; }
    internal PowerPointParagraph? Paragraph { get; }
    internal PowerPointTextBox? List { get; }
    internal static PowerPointTextScope? Current {
        get {
            var scope = OfficeDocumentContext.Find<PowerPointTextScope>();
            return ReferenceEquals(scope?.Owner, PowerPointContentScope.Current) ? scope : null;
        }
    }
}

/// <summary>Maps shared text vocabulary to the native PowerPoint run and theme.</summary>
internal static class PowerPointTextFormatting {
    internal static void Apply(PowerPointTextRun run, PowerPointContentScope scope, bool bold, bool italic,
        bool strikethrough, PowerPointUnderlineStyle? underline, OfficeColor? color, OfficeColor? highlight,
        string? font, OfficeLength? size, string? link = null) {
        if (bold) run.Bold = true;
        if (italic) run.Italic = true;
        if (strikethrough) run.Strikethrough = true;
        if (underline.HasValue) run.UnderlineStyle = underline;
        var theme = new PowerPointLayoutContext(scope.Document, scope.Document.LayoutTheme);
        if (color.HasValue) run.Color = theme.ResolveColor(color);
        if (highlight.HasValue) run.HighlightColor = theme.ResolveColor(highlight);
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
