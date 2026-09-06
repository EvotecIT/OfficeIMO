using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryApplySvgClip(
        OfficeDrawing content, string reference, SvgElementReferenceRegistry references,
        SvgPaintServerRegistry paintServers, OfficeTransform elementTransform,
        double viewX, double viewY, int maximumElements, ref int visited,
        ref int pathCommands, ref bool pathCommandLimitExceeded, ref int unsupported,
        out OfficeDrawing? clipped) {
        clipped = null;
        if (!references.TryEnterLocal(reference, out string id, out XElement? definition)) return false;
        try {
            if (definition == null || !definition.Name.LocalName.Equals("clipPath", StringComparison.OrdinalIgnoreCase)) return false;
            string? nestedClip = ReadPresentationProperty(definition, "clip-path");
            if (!string.IsNullOrWhiteSpace(nestedClip) && !nestedClip!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) return false;
            string? units = definition.Attribute("clipPathUnits")?.Value;
            if (!string.IsNullOrWhiteSpace(units) && !units!.Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) return false;
            XElement[] children = definition.Elements().Where(child =>
                child.Name.LocalName is not "title" and not "desc" and not "metadata").Take(2).ToArray();
            if (children.Length == 0) {
                clipped = new OfficeDrawing(content.Width, content.Height);
                return true;
            }
            // Compound clip unions and referenced/text clip geometry remain diagnosed boundaries.
            if (children.Length != 1 || ++visited > maximumElements) return false;
            XElement child = children[0];
            string? childClip = ReadPresentationProperty(child, "clip-path");
            if (!string.IsNullOrWhiteSpace(childClip) && !childClip!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) return false;
            SvgPaintContext style = ResolveDefinitionPaintContext(child, paintServers, ref unsupported);
            if (!style.Visible || IsEmptySvgClipGeometry(child, content.Width, content.Height)) {
                clipped = new OfficeDrawing(content.Width, content.Height);
                return true;
            }
            OfficeDrawingShape? shape = child.Name.LocalName.ToLowerInvariant() switch {
                "rect" => CreateRectangle(child, style, viewX, viewY, content.Width, content.Height, ref unsupported),
                "circle" => CreateCircle(child, style, viewX, viewY, content.Width, content.Height),
                "ellipse" => CreateEllipse(child, style, viewX, viewY, content.Width, content.Height),
                "polygon" => CreatePolygon(child, style, viewX, viewY, true, ref pathCommands, ref pathCommandLimitExceeded),
                "path" => CreatePath(child, style, viewX, viewY, ref pathCommands, ref pathCommandLimitExceeded),
                _ => null
            };
            if (shape == null) {
                if (child.Name.LocalName.Equals("path", StringComparison.OrdinalIgnoreCase) &&
                    OfficeSvgPathDataParser.TryParse(child.Attribute("d")?.Value, MaximumSvgPathCommands,
                        out var commands, out _, allowEmptyGeometry: true) && commands.All(command => command.Kind is OfficePathCommandKind.MoveTo or OfficePathCommandKind.Close)) {
                    clipped = new OfficeDrawing(content.Width, content.Height);
                    return true;
                }
                return false;
            }
            string? rule = ReadPresentationProperty(child, "clip-rule") ?? ReadPresentationProperty(definition, "clip-rule");
            if (rule != null && !rule.Equals("nonzero", StringComparison.OrdinalIgnoreCase) && !rule.Equals("evenodd", StringComparison.OrdinalIgnoreCase)) return false;
            shape.Shape.FillRule = rule?.Equals("evenodd", StringComparison.OrdinalIgnoreCase) == true
                ? OfficeFillRule.EvenOdd : OfficeFillRule.NonZero;
            OfficeTransform transform = ResolveTransform(definition, elementTransform, viewX, viewY, ref unsupported);
            transform = ResolveTransform(child, transform, viewX, viewY, ref unsupported);
            if (!transform.TryInvert(out _)) {
                clipped = new OfficeDrawing(content.Width, content.Height);
                return true;
            }
            if (!TryCreateDestinationSvgClip(shape, transform, out OfficeClipPath? path, out double x, out double y)) return false;
            clipped = new OfficeDrawing(content.Width, content.Height)
                .AddClippedDrawingForRendering(content, x, y, path!, -x, -y);
            return true;
        } catch (ArgumentException) {
            return false;
        } finally {
            references.Exit(id);
        }
    }
}
