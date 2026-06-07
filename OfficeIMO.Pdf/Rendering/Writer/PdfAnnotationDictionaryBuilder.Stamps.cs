namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationDictionaryBuilder {
    internal static string BuildStampAppearanceContent(double width, double height, string stampName, PdfColor? strokeColor = null, PdfColor? fillColor = null, double borderWidth = 2D) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNullOrWhiteSpace(stampName, nameof(stampName));
        Guard.NonNegative(borderWidth, nameof(borderWidth));

        PdfColor resolvedStrokeColor = strokeColor ?? new PdfColor(0.7D, 0.05D, 0.05D);
        string label = FormatStampLabel(stampName);
        double fontSize = ResolveStampFontSize(label, width, height);
        double textWidth = EstimateWinAnsiTextWidth(label, fontSize);
        double textX = Math.Max(2D, (width - textWidth) / 2D);
        double textY = Math.Max(2D, (height - fontSize) / 2D);

        var builder = new StringBuilder();
        builder.Append("q\n");
        if (fillColor.HasValue) {
            builder.Append(FormatColor(fillColor.Value)).Append(" rg 0 0 ")
                .Append(FormatCoordinate(width)).Append(' ')
                .Append(FormatCoordinate(height)).Append(" re f\n");
        }

        if (borderWidth > 0D) {
            double inset = borderWidth / 2D;
            builder.Append(FormatColor(resolvedStrokeColor)).Append(" RG ")
                .Append(FormatCoordinate(borderWidth)).Append(" w ")
                .Append(FormatCoordinate(inset)).Append(' ')
                .Append(FormatCoordinate(inset)).Append(' ')
                .Append(FormatCoordinate(Math.Max(0D, width - borderWidth))).Append(' ')
                .Append(FormatCoordinate(Math.Max(0D, height - borderWidth))).Append(" re S\n");
        }

        builder.Append("BT /Helv ")
            .Append(FormatCoordinate(fontSize)).Append(" Tf ")
            .Append(FormatColor(resolvedStrokeColor)).Append(" rg ")
            .Append(FormatCoordinate(textX)).Append(' ')
            .Append(FormatCoordinate(textY)).Append(" Td ")
            .Append(PdfSyntaxEscaper.WinAnsiHexString(label))
            .Append(" Tj ET\nQ\n");
        return builder.ToString();
    }

    internal static string BuildCaretAppearanceContent(double width, double height, PdfColor? strokeColor = null, double borderWidth = 1D) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        if (borderWidth <= 0D) {
            return "q\nQ\n";
        }

        PdfColor resolvedStrokeColor = strokeColor ?? PdfColor.Black;
        return "q\n" +
            FormatColor(resolvedStrokeColor) + " RG " +
            FormatCoordinate(borderWidth) + " w 1 J 1 j " +
            "0 " + FormatCoordinate(height) + " m " +
            FormatCoordinate(width / 2D) + " 0 l " +
            FormatCoordinate(width) + " " + FormatCoordinate(height) + " l S\n" +
            "Q\n";
    }

    private static string FormatStampLabel(string stampName) {
        var builder = new StringBuilder();
        for (int i = 0; i < stampName.Length; i++) {
            char ch = stampName[i];
            if (ch == '_' || ch == '-') {
                builder.Append(' ');
                continue;
            }

            if (i > 0 && char.IsUpper(ch) && char.IsLower(stampName[i - 1])) {
                builder.Append(' ');
            }

            builder.Append(char.ToUpperInvariant(ch));
        }

        string label = builder.ToString().Trim();
        return label.Length == 0 ? "STAMP" : label;
    }

    private static double ResolveStampFontSize(string label, double width, double height) {
        double fontSize = Math.Min(18D, Math.Max(7D, height * 0.38D));
        while (fontSize > 7D && EstimateWinAnsiTextWidth(label, fontSize) > Math.Max(1D, width - 8D)) {
            fontSize -= 0.5D;
        }

        return fontSize;
    }
}
