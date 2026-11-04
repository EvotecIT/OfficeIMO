using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static HtmlTextFaceMetrics? ResolveStandardFaceMetrics(PdfCore.PdfCreateOptions options,
        string family, string text, double size, bool bold, bool italic) {
        PdfCore.PdfStandardFont slot = PdfCore.PdfStandardFontMapper.GetStyledFont(MapStandardFont(family), bold, italic);
        if (options.TryGetEmbeddedStandardFontProgram(slot, out PdfCore.PdfTrueTypeFontProgram? trueType)
            && trueType != null && PdfCore.PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, trueType).Count == 0) {
            return FromMetrics(trueType.LineMetrics, trueType.GetAscender(size), trueType.GetDescender(size));
        }
        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(slot, out PdfCore.PdfOpenTypeCffFontProgram? cff)
            && cff != null && PdfCore.PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cff).Count == 0) {
            return FromMetrics(cff.LineMetrics, cff.GetAscender(size), cff.GetDescender(size));
        }
        return null;

        HtmlTextFaceMetrics FromMetrics(OfficeOpenTypeLineMetrics? metrics, double fallbackAscent, double fallbackDescent) {
            double ascent = metrics?.HheaAscentRatio * size ?? fallbackAscent;
            double descent = metrics?.HheaDescentRatio * size ?? fallbackDescent;
            double gap = metrics?.HheaLineGapRatio * size ?? 0D;
            return new HtmlTextFaceMetrics(ascent + descent, ascent, ascent + descent + gap);
        }
    }

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
        PdfCore.PdfCreateOptions options) {
        if (text.Length == 0) return new[] { new NamedFaceStyleRun(text, bold, italic, true) };
        if (NamedFontCoversText(options, family, text, bold, italic)) {
            return new[] { new NamedFaceStyleRun(text, bold, italic, true) };
        }
        if (!options.HasNamedFontFamily(family)) return new[] { new NamedFaceStyleRun(text, bold, italic, false) };

        // Resource policy controls loading new fonts; already registered faces can serve
        // missing glyphs in the same CSS family regardless of their source.
        OfficeFontStyle requested = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        OfficeFontStyle[] styles = [OfficeFontStyle.Regular, OfficeFontStyle.Bold,
            OfficeFontStyle.Italic, OfficeFontStyle.Bold | OfficeFontStyle.Italic];
        Array.Sort(styles, (left, right) => OfficeFontFaceCollection.CompareFaceDescriptors(
            OfficeFontFaceDescriptor.FromStyle(left), OfficeFontFaceDescriptor.FromStyle(right),
            OfficeFontFaceDescriptor.FromStyle(requested)));

        var result = new List<NamedFaceStyleRun>();
        var current = new StringBuilder();
        bool currentBold = bold;
        bool currentItalic = italic;
        bool currentCovered = false;
        foreach (string element in OfficeTextElements.Split(text)) {
            bool selectedBold = bold;
            bool selectedItalic = italic;
            bool covered = false;
            foreach (OfficeFontStyle style in styles) {
                bool candidateBold = (style & OfficeFontStyle.Bold) != 0;
                bool candidateItalic = (style & OfficeFontStyle.Italic) != 0;
                if (!NamedFontCoversText(options, family, element, candidateBold, candidateItalic)) continue;
                selectedBold = candidateBold;
                selectedItalic = candidateItalic;
                covered = true;
                break;
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
        PdfCore.PdfCreateOptions options,
        string family,
        string text,
        bool bold,
        bool italic) =>
        (text.Length > 0 || options.TryResolveNamedFontFace(family, bold, italic, out _))
        && PlanNamedFaceStyleRuns(text, family, bold, italic, options).All(run => run.Covered);

    private static double? MeasureNamedFaceStyledText(
        PdfCore.PdfCreateOptions options,
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
                    fontSize: global::OfficeIMO.OfficeLength.Points(fontSize), fontFamily: family), options);
            if (!measured.HasValue) return null;
            width += measured.Value;
        }
        return width;
    }

    private static HtmlTextFaceMetrics? ResolveNamedFaceMetrics(
        PdfCore.PdfCreateOptions options,
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
        var metricsByFace = new List<HtmlTextFaceMetrics>();
        foreach (NamedFaceStyleRun run in runs) {
            if (!run.Covered || !options.TryResolveNamedFontFace(family, run.Bold, run.Italic,
                    out PdfCore.PdfNamedFontFace face)) return null;
            if (options.TryGetNamedFontProgram(face, out PdfCore.PdfTrueTypeFontProgram? trueType)
                && trueType != null) {
                AddMetrics(trueType.LineMetrics, trueType.GetAscender(fontSize), trueType.GetDescender(fontSize));
            } else if (options.TryGetNamedOpenTypeCffFontProgram(face,
                           out PdfCore.PdfOpenTypeCffFontProgram? cff) && cff != null) {
                AddMetrics(cff.LineMetrics, cff.GetAscender(fontSize), cff.GetDescender(fontSize));
            } else {
                return null;
            }
        }
        return HtmlTextFaceMetrics.Combine(metricsByFace);

        void AddMetrics(OfficeOpenTypeLineMetrics? metrics, double fallbackAscent, double fallbackDescent) {
            double faceAscent = metrics?.HheaAscentRatio * fontSize ?? fallbackAscent;
            double faceDescent = metrics?.HheaDescentRatio * fontSize ?? fallbackDescent;
            double gap = metrics?.HheaLineGapRatio * fontSize ?? 0D;
            metricsByFace.Add(new HtmlTextFaceMetrics(faceAscent + faceDescent, faceAscent,
                faceAscent + faceDescent + gap));
        }
    }
}
