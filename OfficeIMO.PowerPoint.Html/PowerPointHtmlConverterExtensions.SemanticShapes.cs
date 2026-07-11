using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    private static void AppendSemanticShapes(StringBuilder body, PptCore.PowerPointSlide slide, PowerPointHtmlSaveOptions options) {
        foreach (PptCore.PowerPointShape shape in slide.Shapes.OrderBy(shape => shape.DrawingOrder)) {
            if (!options.IncludeHiddenShapes && shape.Hidden) {
                continue;
            }

            if (shape is PptCore.PowerPointTextBox textBox) {
                string text = NormalizeText(textBox.Text);
                if (text.Length == 0) {
                    continue;
                }

                body.Append("<p");
                AppendSemanticShapeAttributes(body, textBox, "text");
                body.Append('>')
                    .Append(OfficeHtmlText.Escape(text))
                    .Append("</p>");
            } else if (shape is PptCore.PowerPointTable table && options.IncludeTables) {
                AppendTable(body, table, includeShapeMetadata: true);
            }
        }
    }

    private static void AppendSemanticShapeAttributes(StringBuilder body, PptCore.PowerPointShape shape, string kind) {
        body.Append(" data-officeimo-layer-kind=\"")
            .Append(kind)
            .Append("\" data-officeimo-layer-index=\"")
            .Append(shape.DrawingOrder.ToString(CultureInfo.InvariantCulture))
            .Append('"');
        AppendDataAttribute(body, "data-officeimo-left", shape.LeftPoints);
        AppendDataAttribute(body, "data-officeimo-top", shape.TopPoints);
        AppendDataAttribute(body, "data-officeimo-width", shape.WidthPoints);
        AppendDataAttribute(body, "data-officeimo-height", shape.HeightPoints);
        AppendDataAttribute(body, "data-officeimo-rotation", shape.Rotation ?? 0D);
        AppendDataAttribute(body, "data-officeimo-flip-horizontal", shape.HorizontalFlip == true);
        AppendDataAttribute(body, "data-officeimo-flip-vertical", shape.VerticalFlip == true);
        AppendDataAttribute(body, "data-officeimo-hidden", shape.Hidden);
    }
}
