using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static bool HasUnsupportedRichText(IWorkTable table, IWorkTableCell cell) {
        if (cell.RichText == null) return false;
        IWorkParagraphStyle? defaults = table.GetParagraphStyle(cell.Row, cell.Column);
        return HasUnsupportedRichText(cell.RichText, cell.Kind == IWorkCellKind.Formula, defaults);
    }

    /// <summary>Separates supported whole-cell formula formatting from unrepresentable result spans and paragraph layout.</summary>
    internal static bool HasUnsupportedRichText(IWorkTextContent content, bool formula,
        IWorkParagraphStyle? defaults = null, bool forceBold = false) =>
        formula && !TryGetUniformFormulaTextStyle(content, forceBold, defaults, out _, out _)
        || content.Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Any(run => run.Hyperlink != null) && UniformCellHyperlink(content) == null
        || RichTextStyles(content).Any(style => style.BackgroundColor != null
            || style.CharacterTrackingEm.GetValueOrDefault() != 0
            || style.Color is { Alpha: < byte.MaxValue })
        || content.Paragraphs.Any(paragraph =>
            paragraph.ListLevel >= 0 || !string.IsNullOrEmpty(paragraph.ListLabel)
            || !formula && paragraph.Style.Alignment.HasValue
            || paragraph.Style.FirstLineIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.LeftIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.RightIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.SpaceBeforePoints.GetValueOrDefault() != 0
            || paragraph.Style.SpaceAfterPoints.GetValueOrDefault() != 0
            || paragraph.Style.LineSpacingMultiplier is double spacing && spacing != 1
            || paragraph.Style.TabStops?.Count > 0
            || paragraph.Style.PageBreakBefore == true
            || paragraph.Style.KeepWithNext == true
            || paragraph.Style.KeepLinesTogether == true
            || paragraph.BreakKind is IWorkParagraphBreakKind.Section
                or IWorkParagraphBreakKind.Layout or IWorkParagraphBreakKind.Page);

    private static IEnumerable<IWorkTextStyle> RichTextStyles(IWorkTextContent content) {
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            if (paragraph.Runs.Count == 0) yield return paragraph.Style.TextStyle;
            else foreach (IWorkTextRun run in paragraph.Runs) yield return run.Style;
        }
    }

    /// <summary>Formula caches cannot store XLSX rich runs; one effective font and alignment can be applied without changing the expression or cache.</summary>
    private static bool TryGetUniformFormulaTextStyle(IWorkTextContent content, bool forceBold,
        IWorkParagraphStyle? defaults, out ExcelRichTextRun? font, out ExcelHorizontalAlignment? alignment) {
        font = null;
        alignment = null;
        if (!content.IsTextComplete || !content.IsFormattingComplete || content.HasUnresolvedInlineObjects) return false;
        bool haveParagraph = false;
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            ExcelHorizontalAlignment? currentAlignment = ResolveFormulaAlignment(paragraph, defaults);
            if (haveParagraph && alignment != currentAlignment) return false;
            alignment = currentAlignment;
            haveParagraph = true;
        }
        foreach (IWorkTextStyle source in RichTextStyles(content)) {
            ExcelRichTextRun current = ToExcelRichTextRun(string.Empty, source, forceBold, defaults?.TextStyle);
            if (font != null && !SameFont(font, current)) return false;
            font = current;
        }
        return font != null;
    }

    private static bool SameFont(ExcelRichTextRun left, ExcelRichTextRun right) =>
        left.Bold == right.Bold && left.Italic == right.Italic
        && left.Underline == right.Underline && left.Strikethrough == right.Strikethrough
        && left.FontSize == right.FontSize
        && string.Equals(left.FontName, right.FontName, StringComparison.Ordinal)
        && string.Equals(left.FontColor, right.FontColor, StringComparison.Ordinal);

    private static ExcelHorizontalAlignment? ResolveFormulaAlignment(IWorkTextParagraph paragraph,
        IWorkParagraphStyle? defaults) => (paragraph.Style.Alignment ?? defaults?.Alignment) switch {
            null => null,
            IWorkTextAlignment.Center => ExcelHorizontalAlignment.Center,
            IWorkTextAlignment.Right => ExcelHorizontalAlignment.Right,
            IWorkTextAlignment.Justified => ExcelHorizontalAlignment.Justify,
            IWorkTextAlignment.Natural => OfficeTextElements.ResolveBaseDirection(paragraph.Text)
                == OfficeTextDirection.RightToLeft ? ExcelHorizontalAlignment.Right : ExcelHorizontalAlignment.General,
            _ => ExcelHorizontalAlignment.Left
        };

    private static void ApplyUniformFormulaTextStyle(ExcelSheet sheet, ExcelCell cell,
        IWorkTextContent content, bool forceBold, IWorkParagraphStyle? defaults) {
        if (!TryGetUniformFormulaTextStyle(content, forceBold, defaults, out ExcelRichTextRun? font,
                out ExcelHorizontalAlignment? alignment)) return;
        cell.SetBold(font!.Bold == true).SetItalic(font.Italic == true)
            .SetUnderline(font.Underline == true).SetStrikethrough(font.Strikethrough == true);
        if (font.FontSize.HasValue) cell.SetFontSize(font.FontSize.Value);
        if (!string.IsNullOrWhiteSpace(font.FontName)) cell.SetFontName(font.FontName!);
        if (font.FontColor != null) cell.SetFontColor(font.FontColor);
        if (alignment.HasValue) sheet.CellAlign(cell.Row, cell.Column, alignment.Value);
        if (UniformCellHyperlink(content) is { } hyperlink)
            sheet.SetHyperlinkReference(cell.Row, cell.Column, hyperlink, style: false);
    }

    private static ExcelRichTextRun ToExcelRichTextRun(string text, IWorkTextStyle style,
        bool forceBold, IWorkTextStyle? defaults) => new(text) {
            Bold = style.Bold ?? defaults?.Bold ?? forceBold,
            Italic = style.Italic ?? defaults?.Italic ?? false,
            Underline = style.Underline ?? defaults?.Underline ?? false,
            Strikethrough = style.Strikethrough ?? defaults?.Strikethrough ?? false,
            FontSize = style.FontSizePoints ?? defaults?.FontSizePoints,
            FontName = style.FontName ?? defaults?.FontName,
            FontColor = style.Color?.RgbHex ?? defaults?.Color?.RgbHex
        };

    private static ExcelRichTextRun[] ToExcelRichTextRuns(IWorkTextContent content, bool forceBold,
        CancellationToken cancellationToken, IWorkTextStyle? defaultStyle = null) {
        var runs = new List<ExcelRichTextRun>();
        for (int paragraphIndex = 0; paragraphIndex < content.Paragraphs.Count; paragraphIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (paragraphIndex > 0) {
                var separator = new ExcelRichTextRun("\n");
                if (forceBold) separator.Bold = true;
                runs.Add(separator);
            }
            foreach (IWorkTextRun source in content.Paragraphs[paragraphIndex].Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                runs.Add(ToExcelRichTextRun(source.Text, source.Style, forceBold, defaultStyle));
            }
        }
        return runs.ToArray();
    }

    private static string? UniformCellHyperlink(IWorkTextContent content) {
        string? hyperlink = null;
        foreach (IWorkTextRun run in content.Paragraphs.SelectMany(paragraph => paragraph.Runs)) {
            if (run.Text.Length == 0) continue;
            if (run.Hyperlink == null) return null;
            if (hyperlink != null && !string.Equals(hyperlink, run.Hyperlink,
                    StringComparison.Ordinal)) return null;
            hyperlink = run.Hyperlink;
        }
        return hyperlink != null && Uri.TryCreate(hyperlink, UriKind.Absolute, out _)
            ? hyperlink : null;
    }
}
