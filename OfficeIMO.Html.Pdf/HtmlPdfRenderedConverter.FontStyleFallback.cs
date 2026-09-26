using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private readonly struct NamedFaceStyleRun {
        internal NamedFaceStyleRun(string text, bool bold, bool italic, bool covered) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Covered = covered;
        }

        internal string Text { get; }
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal bool Covered { get; }
    }

    private static IReadOnlyList<NamedFaceStyleRun> PlanNamedFaceStyleRuns(
        string text,
        string family,
        bool bold,
        bool italic,
        PdfCore.PdfOptions options) {
        if (text.Length == 0) return new[] { new NamedFaceStyleRun(text, bold, italic, true) };
        if (NamedFontCoversText(options, family, text, bold, italic)) {
            return new[] { new NamedFaceStyleRun(text, bold, italic, true) };
        }
        if (!bold && !italic) return new[] { new NamedFaceStyleRun(text, bold, italic, false) };

        var result = new List<NamedFaceStyleRun>();
        var current = new StringBuilder();
        bool currentBold = bold;
        bool currentItalic = italic;
        bool currentCovered = false;
        foreach (string element in OfficeTextElements.Split(text)) {
            bool selectedBold = bold;
            bool selectedItalic = italic;
            bool covered = NamedFontCoversText(options, family, element, bold, italic);
            if (!covered && bold && italic && NamedFontCoversText(options, family, element, true, false)) {
                selectedItalic = false;
                covered = true;
            }
            if (!covered && bold && italic && NamedFontCoversText(options, family, element, false, true)) {
                selectedBold = false;
                covered = true;
            }
            if (!covered && NamedFontCoversText(options, family, element, false, false)) {
                selectedBold = false;
                selectedItalic = false;
                covered = true;
            }
            if (current.Length > 0 && (currentBold != selectedBold || currentItalic != selectedItalic || currentCovered != covered)) {
                result.Add(new NamedFaceStyleRun(current.ToString(), currentBold, currentItalic, currentCovered));
                current.Clear();
            }
            currentBold = selectedBold;
            currentItalic = selectedItalic;
            currentCovered = covered;
            current.Append(element);
        }
        if (current.Length > 0) result.Add(new NamedFaceStyleRun(current.ToString(), currentBold, currentItalic, currentCovered));
        return result;
    }

    private static bool NamedFontCoversTextWithStyleFallback(
        PdfCore.PdfOptions options,
        string family,
        string text,
        bool bold,
        bool italic) =>
        PlanNamedFaceStyleRuns(text, family, bold, italic, options).All(run => run.Covered);

    private static double? MeasureNamedFaceStyledText(
        PdfCore.PdfOptions options,
        string family,
        string text,
        double fontSize,
        bool bold,
        bool italic) {
        if (!options.HasNamedFontFamily(family)) return null;
        double width = 0D;
        foreach (NamedFaceStyleRun run in PlanNamedFaceStyleRuns(text, family, bold, italic, options)) {
            double? measured = PdfCore.PdfWriter.MeasurePositionedText(
                new PdfCore.PdfTextRun(run.Text, bold: run.Bold, italic: run.Italic,
                    fontSize: fontSize, fontFamily: family), options);
            if (!measured.HasValue) return null;
            width += measured.Value;
        }
        return width;
    }

    private static HtmlTextFaceMetrics? ResolveNamedFaceMetrics(
        PdfCore.PdfOptions options,
        string family,
        string text,
        double fontSize,
        bool bold,
        bool italic,
        bool allowStyleFallback) {
        IReadOnlyList<NamedFaceStyleRun> runs = allowStyleFallback
            ? PlanNamedFaceStyleRuns(text, family, bold, italic, options)
            : new[] { new NamedFaceStyleRun(text, bold, italic,
                NamedFontCoversText(options, family, text, bold, italic)) };
        double ascent = 0D;
        double descent = 0D;
        foreach (NamedFaceStyleRun run in runs) {
            if (!run.Covered || !options.TryResolveNamedFontFace(family, run.Bold, run.Italic,
                    out PdfCore.PdfNamedFontFace face)) return null;
            if (options.TryGetNamedFontProgram(face, out PdfCore.PdfTrueTypeFontProgram? trueType)
                && trueType != null) {
                ascent = System.Math.Max(ascent, trueType.GetAscender(fontSize));
                descent = System.Math.Max(descent, trueType.GetDescender(fontSize));
            } else if (options.TryGetNamedOpenTypeCffFontProgram(face,
                           out PdfCore.PdfOpenTypeCffFontProgram? cff) && cff != null) {
                ascent = System.Math.Max(ascent, cff.GetAscender(fontSize));
                descent = System.Math.Max(descent, cff.GetDescender(fontSize));
            } else {
                return null;
            }
        }
        return ascent > 0D ? new HtmlTextFaceMetrics(ascent + descent, ascent) : null;
    }
}
