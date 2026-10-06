using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static bool HasSelectedVariableFont(OfficeDrawingText text, OfficeFontFaceCollection fonts) {
        double size = (text.Font.Size > 0D ? text.Font.Size : 10D) / text.FontMetricScale;
        foreach (OfficeFontFallbackRun run in fonts.PlanFallbackRuns(text.RasterText, text.Font.FamilyName, text.Font.Face)) {
            if (fonts.TryResolveFaceForText(run.Text, run.FamilyName, text.Font.Face, size, out OfficeFontFace? face)
                && face?.VariationCoordinatesForShaping is { Count: > 0 }) return true;
        }
        return false;
    }

    // Process only renderer-produced text markup. Keep line positioning, clipping and
    // decoration elements intact, and bind variation coordinates to each selected face.
    private static void AppendSelectedVariableFonts(StringBuilder output, string fragment,
        OfficeDrawingText text, OfficeFontFaceCollection fonts, System.Threading.CancellationToken cancellationToken) {
        XElement root = XElement.Parse("<g>" + fragment + "</g>", LoadOptions.PreserveWhitespace);
        foreach (XText node in root.DescendantNodes().OfType<XText>().ToArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node.Parent?.Name.LocalName is not ("text" or "tspan")) continue;
            string family = node.Parent.AncestorsAndSelf().Select(e => e.Attribute("font-family")?.Value)
                .FirstOrDefault(value => value != null) ?? text.Font.FamilyName ?? "Arial";
            string? sizeValue = node.Parent.AncestorsAndSelf().Select(e => e.Attribute("font-size")?.Value)
                .FirstOrDefault(value => value != null);
            double size = double.TryParse(sizeValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double renderedSize)
                && renderedSize > 0D ? renderedSize : text.Font.Size;
            var spans = fonts.PlanFallbackRuns(node.Value, family, text.Font.Face).Select(run => {
                fonts.TryResolveFaceForText(run.Text, run.FamilyName, text.Font.Face,
                    size / text.FontMetricScale, out OfficeFontFace? face);
                var span = new XElement("tspan", new XAttribute("font-family",
                    face == null ? run.FamilyName : QuoteCssFamily(face.ResourceFamilyName)), run.Text);
                if (face?.VariationCoordinatesForShaping is { Count: > 0 } coordinates) {
                    string values = string.Join(",", coordinates.OrderBy(item => item.Key, StringComparer.Ordinal)
                        .Select(item => "\"" + EscapeCssString(item.Key) + "\" "
                            + item.Value.ToString("R", CultureInfo.InvariantCulture)));
                    span.SetAttributeValue("style", "font-optical-sizing:none;font-variation-settings:" + values);
                }
                return span;
            }).ToArray();
            node.ReplaceWith(spans);
        }
        foreach (XNode node in root.Nodes()) output.Append(node.ToString(SaveOptions.DisableFormatting));
    }
}
