using System.Xml;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static partial class VisioSvgRenderer {
    private static void WriteRichText(XmlWriter writer, VisioRichTextProjection projection,
        double x, double y, VisioTextStyle? style, double rotationRadians, double width, double height,
        bool drawLabelBackground, bool labelAdjusted, CancellationToken cancellationToken) {
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Arial", 10D));
        OfficeRichTextBlockLayout layout = VisioRichTextLayout.Create(projection, width, height,
            (text, size, family, fontStyle) => MeasureSvgTextWidth(measurer, text, size, family ?? "Arial", fontStyle),
            5D, cancellationToken);
        writer.WriteStartElement("g", SvgNamespace);
        writer.WriteAttributeString("data-officeimo-rich-text", "true");
        ConfigureSvgTextAttributes(writer, style, labelAdjusted);
        if (ResolveTextBackground(style, drawLabelBackground) is OfficeColor background && background.A > 0) {
            OfficeTextBlockBackgroundBounds bounds = VisioRichTextLayout.Background(layout, projection, style,
                x, y, width, height, 3D, 2D);
            writer.WriteStartElement("rect", SvgNamespace);
            writer.WriteAttributeString("x", OfficeSvgFormatting.FormatNumber(bounds.Left));
            writer.WriteAttributeString("y", OfficeSvgFormatting.FormatNumber(bounds.Top));
            writer.WriteAttributeString("width", OfficeSvgFormatting.FormatNumber(bounds.Width));
            writer.WriteAttributeString("height", OfficeSvgFormatting.FormatNumber(bounds.Height));
            OfficeSvgFormatting.WriteColorAttribute(writer, "fill", background);
            if (Math.Abs(rotationRadians) > 0.000001D)
                writer.WriteAttributeString("transform", FormatTextRotation(rotationRadians, x, y));
            ConfigureTextBackgroundAttributes(writer, drawLabelBackground, labelAdjusted);
            writer.WriteEndElement();
        }
        var markup = new StringBuilder();
        markup.AppendSvgRichTextBlock(layout, x - width / 2D, y - height / 2D, width, height,
            projection.RenderAlignment, VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment),
            RadiansToDegrees(-rotationRadians), x, y, centerLineInLineHeight: false);
        writer.WriteRaw(markup.ToString());
        writer.WriteEndElement();
    }
}
