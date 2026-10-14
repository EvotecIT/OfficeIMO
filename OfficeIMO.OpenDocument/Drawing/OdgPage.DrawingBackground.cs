using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static readonly HashSet<XName> BackgroundProperties = new HashSet<XName> {
        OdfNamespaces.Draw + "fill", OdfNamespaces.Draw + "fill-color", OdfNamespaces.Draw + "background-size",
        OdfNamespaces.Draw + "fill-image-name", OdfNamespaces.Style + "repeat", OdfNamespaces.Draw + "opacity",
        OdfNamespaces.Draw + "opacity-name", OdfNamespaces.Draw + "fill-gradient-name", OdfNamespaces.Draw + "fill-hatch-name",
        OdfNamespaces.Draw + "fill-hatch-solid", OdfNamespaces.Draw + "gradient-step-count", OdfNamespaces.Draw + "secondary-fill-color",
        OdfNamespaces.Draw + "fill-image-width", OdfNamespaces.Draw + "fill-image-height", OdfNamespaces.Draw + "fill-image-ref-point",
        OdfNamespaces.Draw + "fill-image-ref-point-x", OdfNamespaces.Draw + "fill-image-ref-point-y", OdfNamespaces.Draw + "tile-repeat-offset"
    };

    private void ProjectBackground(OfficeDrawing drawing, OdfConversionReport report) {
        XElement[] page = DrawingPageProperties(Element, "content.xml", report, "page-style");
        XElement[] master = Master is XElement source ? DrawingPageProperties(source, "styles.xml", report, "master-page-style") : Array.Empty<XElement>();
        // Native Draw uses the master's paint area, and a page no-fill keeps the master fill.
        // Presentation-only visibility declarations retain separate loss diagnostics.
        string? pageFill = Property(page, OdfNamespaces.Draw + "fill");
        XElement[] selected = pageFill is null or "none" ? master : page;
        string? fill = Property(selected, OdfNamespaces.Draw + "fill");
        if (fill is null or "none") return;
        const string feature = "page-background";
        try {
            if (fill is not ("solid" or "bitmap" or "gradient")) throw new NotSupportedException("Only solid, supported native two-color gradient and stretched embedded bitmap page backgrounds are projected.");
            if (!string.IsNullOrEmpty(Property(selected, OdfNamespaces.Draw + "opacity-name")))
                throw new NotSupportedException("Background opacity gradients are preserved but not projected.");
            string? alpha = Property(selected, OdfNamespaces.Draw + "opacity");
            double opacity = alpha == null ? 1 : OdfOpacity.Parse(alpha);
            string size = Property(master, OdfNamespaces.Draw + "background-size") ?? "border";
            string? pageSize = Property(page, OdfNamespaces.Draw + "background-size");
            if (pageSize != null && pageSize != size)
                report.Add(feature + ":page-area", OdfConversionMappingStatus.Unsupported,
                    message: "The page background-size declaration is preserved; this Draw projection uses the master's paint area, as in the qualified native controls.");
            OfficeImagePlacement area = BackgroundArea(drawing, size);
            OdfGradientPattern? gradient = null;
            if (fill is "solid" or "gradient") {
                var shape = OfficeShape.Rectangle(area.Width, area.Height);
                shape.FillColor = null; shape.StrokeColor = null; shape.FillOpacity = opacity;
                if (fill == "solid") {
                    string color = Property(selected, OdfNamespaces.Draw + "fill-color")
                        ?? throw new InvalidDataException("A solid background requires a resolved fill color.");
                    shape.FillColor = OfficeColor.Parse(OdfColor.Parse(color).ToString());
                } else {
                    string name = Property(selected, OdfNamespaces.Draw + "fill-gradient-name")
                        ?? throw new InvalidDataException("A gradient background requires a named definition.");
                    gradient = (_document.Styles.FindGradient(name) ?? throw new InvalidDataException("Missing background gradient '" + name + "'.")).Pattern;
                    ApplyNativeGradientFill(shape, gradient, BackgroundGradientSteps(selected));
                }
                drawing.AddShape(shape, area.X, area.Y);
            } else {
                if (Property(selected, OdfNamespaces.Style + "repeat") != "stretch")
                    throw new NotSupportedException("Tiled, unstretched and unspecified bitmap repeat modes are preserved but not projected.");
                // ODF ignores source size and tile alignment/offset declarations in stretch mode.
                string name = Property(selected, OdfNamespaces.Draw + "fill-image-name")
                    ?? throw new InvalidDataException("A bitmap background requires a named fill image.");
                byte[] bytes = _document.Styles.ReadFillImage(name);
                if (!OfficeImageReader.TryValidateContent(bytes, null, out OfficeImageInfo image) || !IsWithinRasterProjectionProfile(image))
                    throw new NotSupportedException("The embedded background is unavailable or outside the supported raster image profile.");
                drawing.AddImageWithInterpolation(bytes, image.MimeType, new OfficeImageProjection(area), interpolate: false, opacity: opacity);
            }
            report.Add(feature, gradient == null ? OdfConversionMappingStatus.Converted : OdfConversionMappingStatus.Approximated,
                message: "The resolved " + fill + " background paints beneath master and page artwork using the master's " + size + " area." +
                    (gradient == null ? "" : " " + NativeGradientMessage(gradient)));
        } catch (Exception exception) when (exception is ArgumentException or FormatException or InvalidDataException or NotSupportedException or OverflowException) {
            report.Add(feature, OdfConversionMappingStatus.Skipped, message: "Native background is preserved but not projected: " + exception.Message);
        }
    }

    private static int? BackgroundGradientSteps(IEnumerable<XElement> properties) {
        string? text = Property(properties, OdfNamespaces.Draw + "gradient-step-count");
        if (text == null) return null;
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) || count < 0 || count is 1 or 2)
            throw new InvalidDataException("Invalid background gradient step count.");
        return count;
    }

    private XElement[] DrawingPageProperties(XElement owner, string part, OdfConversionReport report, string feature) {
        string? name = (string?)owner.Attribute(OdfNamespaces.Draw + "style-name");
        OdfStyle? style = name == null ? null : _document.Styles.FindInPart(OdfStyleFamily.DrawingPage, name, part);
        XElement[] properties = _document.Styles.ResolveWithDefault(style, OdfStyleFamily.DrawingPage)
            .Select(candidate => candidate.Element.Element(OdfNamespaces.Style + "drawing-page-properties")).OfType<XElement>().ToArray();
        if (name != null && style == null || properties.Any(element => element.HasElements || element.Attributes().Any(attribute =>
            !attribute.IsNamespaceDeclaration && !BackgroundProperties.Contains(attribute.Name))))
            report.Add(feature, OdfConversionMappingStatus.Unsupported,
                message: "Unresolved drawing-page styles and presentation effects are preserved but are outside the Draw projection profile.");
        return properties;
    }

    private OfficeImagePlacement BackgroundArea(OfficeDrawing drawing, string size) {
        if (size == "full") return new OfficeImagePlacement(0, 0, drawing.Width, drawing.Height);
        if (size != "border") throw new NotSupportedException("Unknown background-size declaration.");
        XElement properties = LayoutProperties;
        double left = Margin("left"), right = Margin("right"), top = Margin("top"), bottom = Margin("bottom");
        double width = drawing.Width - left - right, height = drawing.Height - top - bottom;
        if (!(width > 0) || !(height > 0)) throw new NotSupportedException("Background margins require a positive paint area inside the page.");
        return new OfficeImagePlacement(left, top, width, height);
        double Margin(string side) {
            string? value = (string?)properties.Attribute(OdfNamespaces.Fo + "margin-" + side) ?? (string?)properties.Attribute(OdfNamespaces.Fo + "margin");
            if (value == null) return 0;
            if (!OdfLength.Parse(value).TryToPoints(out double points) || points < 0)
                throw new NotSupportedException("Background margins require finite nonnegative absolute lengths.");
            return points;
        }
    }

    private static string? Property(IEnumerable<XElement> properties, XName name) =>
        properties.Select(element => (string?)element.Attribute(name)).FirstOrDefault(value => value != null);
}
