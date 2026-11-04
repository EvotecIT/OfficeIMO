using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetColumnFontMetrics(
            ExcelDocument document,
            PdfCore.PdfCreateOptions pdfOptions,
            ExcelPdfExportOptions options,
            IReadOnlyList<WorksheetPdfExportPlan> plans) {
            if (options.WorksheetLayout != ExcelPdfWorksheetLayoutMode.WorksheetCanvas ||
                !options.UseWorksheetColumnWidths ||
                document.GetWorkbookDefaultFontInfo() is not OfficeFontInfo defaultFont ||
                !pdfOptions.TryResolveNamedFontFace(defaultFont.FamilyName, defaultFont.IsBold, defaultFont.IsItalic,
                    out PdfCore.PdfNamedFontFace face, allowSubstitution: false)) return;

            Func<string, double> measure;
            Func<int, bool> covers;
            if (pdfOptions.TryGetNamedFontProgram(face, out PdfCore.PdfTrueTypeFontProgram? trueType) && trueType != null) {
                measure = text => trueType.MeasureTextWidth(text, defaultFont.Size);
                covers = scalar => trueType.TryGetGlyphId(scalar, out int glyph) && glyph > 0;
            } else if (pdfOptions.TryGetNamedOpenTypeCffFontProgram(face, out PdfCore.PdfOpenTypeCffFontProgram? cff) && cff != null) {
                measure = text => cff.MeasureTextWidth(text, defaultFont.Size);
                covers = scalar => cff.TryGetGlyphId(scalar, out int glyph) && glyph > 0;
            } else {
                return;
            }

            double maximum = 0D;
            for (char digit = '0'; digit <= '9'; digit++) {
                if (!covers(digit)) return;
                double width = measure(digit.ToString());
                if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) return;
                maximum = Math.Max(maximum, width);
            }

            foreach (WorksheetPdfExportPlan plan in plans) plan.Geometry.MaximumDigitWidthPoints = maximum;
        }
    }
}
