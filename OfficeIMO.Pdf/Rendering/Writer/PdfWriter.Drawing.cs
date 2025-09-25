namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string SetFillColor(PdfColor color) => F(color.R) + " " + F(color.G) + " " + F(color.B) + " rg\n";
    private static string SetStrokeColor(PdfColor color) => F(color.R) + " " + F(color.G) + " " + F(color.B) + " RG\n";

    private static void DrawRowFill(StringBuilder sb, PdfColor color, double x, double y, double w, double h) {
        sb.Append("q\n");
        sb.Append(SetFillColor(color));
        sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re f\n");
        sb.Append("Q\n");
    }

    private static void DrawRowRect(StringBuilder sb, PdfColor color, double widthStroke, double x, double y, double w, double h) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re S\n");
        sb.Append("Q\n");
    }

    private static void DrawVLine(StringBuilder sb, PdfColor color, double widthStroke, double x, double yTop, double yBottom) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x)).Append(' ').Append(F(yTop)).Append(" m ").Append(F(x)).Append(' ').Append(F(yBottom)).Append(" l S\n");
        sb.Append("Q\n");
    }

    private static void DrawHLine(StringBuilder sb, PdfColor color, double widthStroke, double x1, double x2, double y) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x1)).Append(' ').Append(F(y)).Append(" m ").Append(F(x2)).Append(' ').Append(F(y)).Append(" l S\n");
        sb.Append("Q\n");
    }

    private static void WriteCell(StringBuilder sb, string fontRes, double fontSize, double x, double y, string text, PdfColor? color, PdfOptions opts) {
        sb.Append("BT\n");
        sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
        var effective = color ?? opts.DefaultTextColor;
        if (effective.HasValue) sb.Append(SetFillColor(effective.Value));
        sb.Append("1 0 0 1 ").Append(F(x)).Append(' ').Append(F(y)).Append(" Tm\n");
        sb.Append('<').Append(EncodeWinAnsiHex(text)).Append("> Tj\n");
        sb.Append("ET\n");
    }
}

