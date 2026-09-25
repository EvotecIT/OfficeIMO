using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static int CountUnmappedPowerPointShapeAppearance(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            if (part.Slide == null) continue;
            count += part.Slide.Descendants<P.ShapeProperties>().Count(HasUnmappedPowerPointShapeAppearance);
            count += part.Slide.Descendants<P.ShapeStyle>().Count();
            count += part.Slide.Descendants<P.Shape>().Count(shape =>
                shape.TextBody?.BodyProperties is A.BodyProperties body && (body.HasAttributes || body.HasChildren));
            count += part.Slide.Descendants<P.Picture>().Count(HasUnmappedPictureBlipAppearance);
        }
        return count;
    }

    private static int CountUnmappedPowerPointNonTextPlaceholders(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += part.Slide.Descendants<P.Picture>().Count(picture =>
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.GraphicFrame>().Count(frame =>
                frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.ConnectionShape>().Count(connection =>
                connection.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.Shape>().Count(shape =>
                shape.TextBody == null && shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() != null);
        }
        return count;
    }

    private static int CountUnmappedPowerPointShapeAccessibility(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += part.Slide.Descendants<P.NonVisualDrawingProperties>().Count(properties =>
                properties.GetAttributes().Any(attribute => attribute.LocalName is "title" or "descr") ||
                properties.Descendants().Any(child => child.LocalName == "decorative"));
        }
        return count;
    }

    private static bool HasUnmappedPowerPointShapeAppearance(P.ShapeProperties properties) {
        OpenXmlElement[] fills = properties.ChildElements.Where(IsFillElement).ToArray();
        if (fills.Length > 1 || fills.Length == 1 && !IsDirectRgbFill(fills[0])) return true;
        A.Outline? outline = properties.GetFirstChild<A.Outline>();
        if (outline != null && !IsDirectRgbOutline(outline)) return true;
        return properties.ChildElements.Any(child => child.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d");
    }

    private static bool IsFillElement(OpenXmlElement element) => element is
        A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill or A.NoFill;

    private static bool IsDirectRgbFill(OpenXmlElement element) =>
        element is A.SolidFill solid && solid.ChildElements.Count == 1 &&
        solid.RgbColorModelHex is { ChildElements.Count: 0 };

    private static bool IsDirectRgbOutline(A.Outline outline) {
        if (outline.GetAttributes().Any(attribute => attribute.LocalName != "w")) return false;
        return outline.ChildElements.Count == 1 && IsDirectRgbFill(outline.ChildElements[0]);
    }

    private static bool HasUnmappedPictureBlipAppearance(P.Picture picture) {
        P.BlipFill? fill = picture.BlipFill;
        if (fill == null) return false;
        A.Blip? blip = fill.Blip;
        return blip?.ChildElements.Count > 0 ||
            fill.ChildElements.Any(child => child is not A.Blip and not A.SourceRectangle and not A.Stretch) ||
            fill.Descendants<A.FillRectangle>().Any(rectangle => rectangle.HasAttributes);
    }

    private static int CountUnmappedPowerPointTextGeometry(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            count += part.Slide?.Descendants<P.Shape>().Count(shape => {
                if (shape.TextBody == null) return false;
                P.ShapeProperties? properties = shape.ShapeProperties;
                A.PresetGeometry? preset = properties?.GetFirstChild<A.PresetGeometry>();
                return properties?.GetFirstChild<A.CustomGeometry>() != null ||
                    preset != null && (preset.Preset?.Value != A.ShapeTypeValues.Rectangle ||
                                       preset.AdjustValueList?.HasChildren == true);
            }) ?? 0;
        }
        return count;
    }

    private static int CountUnwrappedOdpDrawingElements(OdpSlide slide) {
        var wrapped = new HashSet<XElement>(slide.Shapes.Select(shape => shape.Element));
        return slide.Element.Elements().Count(element =>
            element.Name.Namespace == OdfNamespaces.Draw &&
            element.Name != OdfNamespaces.Draw + "page-thumbnail" && !wrapped.Contains(element));
    }

    private static bool HasUnmappedOdpShapeAppearance(OdpPresentation source, OdpShape shape) {
        string? styleName = (string?)shape.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.Graphic, styleName,
            OdfNamespaces.Style + "graphic-properties");
        if (properties == null) return false;
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        string? stroke = (string?)properties.Attribute(OdfNamespaces.Draw + "stroke");
        if (fill != null && fill != "solid" || stroke != null && stroke != "solid") return true;
        if (fill == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out _)) return true;
        if (stroke == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Svg + "stroke-color"), out _)) return true;
        return properties.HasElements || properties.Attributes().Any(attribute =>
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color" &&
            attribute.Name != OdfNamespaces.Draw + "stroke" && attribute.Name != OdfNamespaces.Svg + "stroke-color" &&
            attribute.Name != OdfNamespaces.Svg + "stroke-width");
    }

    private static bool HasUnmappedOdpNoteContent(OdpPresentation source, OdpSlide slide) {
        XElement? notes = slide.Element.Element(OdfNamespaces.Presentation + "notes");
        if (notes == null) return false;
        bool defaultFrameAppearance = notes.Descendants(OdfNamespaces.Draw + "frame").Any() &&
            source.Package.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "default-style")
                .Any(candidate => (string?)candidate.Attribute(OdfNamespaces.Style + "family") == "graphic" &&
                    candidate.Element(OdfNamespaces.Style + "graphic-properties") is XElement properties &&
                    (properties.HasAttributes || properties.HasElements));
        return notes.HasAttributes || defaultFrameAppearance || notes.Descendants().Any(element =>
            element.Name.Namespace == OdfNamespaces.Table ||
            element.Name.Namespace == OdfNamespaces.Draw &&
            element.Name != OdfNamespaces.Draw + "frame" &&
            element.Name != OdfNamespaces.Draw + "text-box" &&
            element.Name != OdfNamespaces.Draw + "page-thumbnail" ||
            element.Name == OdfNamespaces.Draw + "frame" && element.Attributes().Any(attribute =>
                attribute.Name != OdfNamespaces.Draw + "name" &&
                (attribute.Name != OdfNamespaces.Presentation + "class" || attribute.Value != "notes")) ||
            element.Name == OdfNamespaces.Draw + "text-box" && element.HasAttributes);
    }

    private static bool HasUnmappedOdpTransitionTiming(OdpPresentation source, OdpSlide slide) {
        string? styleName = (string?)slide.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage, styleName,
            OdfNamespaces.Style + "drawing-page-properties");
        return properties?.Attribute(OdfNamespaces.Presentation + "transition-change") != null ||
            properties?.Attribute(OdfNamespaces.Presentation + "duration") != null ||
            slide.Element.Attribute(OdfNamespaces.Presentation + "transition-change") != null ||
            slide.Element.Attribute(OdfNamespaces.Presentation + "duration") != null;
    }

    private static int CountUnmappedPowerPointTableAppearance(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            if (part.Slide != null)
                count += part.Slide.Descendants<A.Table>().Count(HasUnmappedPowerPointTableAppearance);
        }
        return count;
    }

    private static bool HasUnmappedPowerPointTableAppearance(A.Table table) {
        A.TableProperties? properties = table.TableProperties;
        if (properties != null && (properties.HasAttributes || properties.HasChildren)) return true;
        long[] columnWidths = table.TableGrid?.Elements<A.GridColumn>()
            .Select(column => column.Width?.Value ?? 0L).ToArray() ?? Array.Empty<long>();
        long[] rowHeights = table.Elements<A.TableRow>()
            .Select(row => row.Height?.Value ?? 0L).ToArray();
        if (columnWidths.Distinct().Skip(1).Any() || rowHeights.Distinct().Skip(1).Any()) return true;
        A.Extents? extents = table.Ancestors<P.GraphicFrame>().FirstOrDefault()?.Transform?.Extents;
        if (extents != null && (columnWidths.Sum() != extents.Cx?.Value ||
            rowHeights.Sum() != extents.Cy?.Value)) return true;
        return table.Descendants<A.TableCellProperties>().Any(cell => cell.HasAttributes || cell.HasChildren) ||
            table.Descendants<A.BodyProperties>().Any(body => body.HasAttributes || body.HasChildren);
    }

    private static bool HasUnmappedOdpTableAppearance(OdpTable table) =>
        table.Element.DescendantsAndSelf().Any(element => element.Attributes().Any(attribute =>
            attribute.Name == OdfNamespaces.Table + "style-name" ||
            attribute.Name == OdfNamespaces.Table + "default-cell-style-name"));

    private static bool HasUnmappedOdpTableValues(OdpTable table) =>
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Any(cell =>
            cell.Attributes().Any(attribute => attribute.Name.Namespace == OdfNamespaces.Office &&
                attribute.Name.LocalName is "value-type" or "value" or "date-value" or "time-value" or
                    "boolean-value" or "string-value" or "currency"));

    private static int CountUnmappedOdpCustomShows(OdpPresentation source) =>
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Presentation + "show").Count();

    private static (bool Override, bool Loss, OdfColor? Color, bool SuppressesMasterBackground) ReadOdpSlideBackground(
        OdpPresentation source, OdpSlide slide) {
        string? styleName = (string?)slide.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage, styleName,
            OdfNamespaces.Style + "drawing-page-properties");
        bool suppressesMasterBackground = string.Equals(
            (string?)slide.Element.Attribute(OdfNamespaces.Presentation + "background-visible"),
            "false", StringComparison.OrdinalIgnoreCase) || string.Equals(
            (string?)properties?.Attribute(OdfNamespaces.Presentation + "background-visible"),
            "false", StringComparison.OrdinalIgnoreCase);
        if (properties == null) return (false, false, null, suppressesMasterBackground);
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        if (fill == null) {
            bool unsupportedInheritedProperties = properties.Attributes().Any(attribute =>
                attribute.Name.Namespace == OdfNamespaces.Draw &&
                attribute.Name != OdfNamespaces.Draw + "fill-color");
            return (false, unsupportedInheritedProperties, null, suppressesMasterBackground);
        }
        OdfColor? color = fill == "solid" &&
            OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out OdfColor parsed)
            ? parsed : (OdfColor?)null;
        bool loss = fill != "none" && !color.HasValue || properties.Attributes().Any(attribute =>
            attribute.Name.Namespace == OdfNamespaces.Draw &&
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color");
        return (true, loss, color, suppressesMasterBackground);
    }

    private static XElement? EffectiveOdfStyleProperties(OdpPresentation source, OdfStyleFamily family,
        string? styleName, XName propertiesName, string partPath = "content.xml") {
        var effective = new XElement(propertiesName);
        OdfStyle? style = string.IsNullOrWhiteSpace(styleName) ? null :
            source.Styles.FindInPart(family, styleName!, partPath);
        foreach (OdfStyle candidate in style == null ? Array.Empty<OdfStyle>() : source.Styles.Resolve(style)) {
            XElement? properties = candidate.Element.Element(propertiesName);
            if (properties == null) continue;
            foreach (XAttribute attribute in properties.Attributes()) {
                if (effective.Attribute(attribute.Name) == null)
                    effective.SetAttributeValue(attribute.Name, attribute.Value);
            }
            foreach (XElement child in properties.Elements()) effective.Add(new XElement(child));
        }
        string familyName = family == OdfStyleFamily.DrawingPage ? "drawing-page" :
            family.ToString().ToLowerInvariant();
        XElement? defaultProperties = source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "default-style")
            .FirstOrDefault(candidate => (string?)candidate.Attribute(OdfNamespaces.Style + "family") ==
                familyName)?.Element(propertiesName);
        if (defaultProperties != null) {
            foreach (XAttribute attribute in defaultProperties.Attributes()) {
                if (effective.Attribute(attribute.Name) == null)
                    effective.SetAttributeValue(attribute.Name, attribute.Value);
            }
            foreach (XElement child in defaultProperties.Elements()) {
                if (!effective.Elements(child.Name).Any()) effective.Add(new XElement(child));
            }
        }
        if ((string?)effective.Attribute(OdfNamespaces.Draw + "fill") == "solid") {
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-image-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-hatch-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-transparency-gradient-name", null);
        }
        return effective.HasAttributes || effective.HasElements ? effective : null;
    }
}
