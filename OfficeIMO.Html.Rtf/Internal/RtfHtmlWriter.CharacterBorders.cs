using System.Text;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendCharacterBorderStyle(StringBuilder builder, RtfCharacterBorder border, RtfDocument document) {
        if (!border.HasAnyValue) {
            return;
        }

        builder.Append("border:");
        if (border.Width.HasValue) {
            builder.Append(FormatPoints(border.Width.Value / 20d));
            builder.Append("pt ");
        }

        builder.Append(FormatParagraphBorderStyle(border.Style));
        if (TryGetColor(document, border.ColorIndex, out RtfColor? color)) {
            builder.Append(' ');
            builder.Append(FormatColor(color!));
        }

        builder.Append(';');
    }
}
