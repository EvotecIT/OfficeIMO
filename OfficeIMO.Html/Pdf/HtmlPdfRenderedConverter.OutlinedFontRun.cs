using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private readonly struct OutlinedFontRun {
        internal OutlinedFontRun(string text, OfficeFontFace face, bool simulateBold, bool simulateItalic)
            : this(text, face, simulateBold, simulateItalic, 0D, null, null) {
        }

        internal OutlinedFontRun(
            string text,
            OfficeFontFace face,
            bool simulateBold,
            bool simulateItalic,
            double advance,
            string? shapedText,
            OfficeTextShapingResult? shapingResult) {
            Text = text;
            Face = face;
            SimulateBold = simulateBold;
            SimulateItalic = simulateItalic;
            Advance = advance;
            ShapedText = shapedText;
            ShapingResult = shapingResult;
            PreparedShaping = null;
        }

        internal OutlinedFontRun(OfficePreparedTextShaping prepared, double advance) {
            Text = prepared.Text;
            Face = null;
            PreparedShaping = prepared;
            SimulateBold = false;
            SimulateItalic = false;
            Advance = advance;
            ShapedText = prepared.Text;
            ShapingResult = prepared.Result;
        }

        internal string Text { get; }
        internal OfficeFontFace? Face { get; }
        internal OfficePreparedTextShaping? PreparedShaping { get; }
        internal IOfficeFontProgram Program => PreparedShaping?.Program ?? Face!.Program;
        internal bool SimulateBold { get; }
        internal bool SimulateItalic { get; }
        internal double Advance { get; }
        internal string? ShapedText { get; }
        internal OfficeTextShapingResult? ShapingResult { get; }
    }

}
