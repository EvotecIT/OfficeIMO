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

    private static bool HasUnmappedOdpShapeAppearance(OdpPresentation source, OdpShape shape) {
        string? styleName = (string?)shape.Element.Attribute(OdfNamespaces.Draw + "style-name");
        if (styleName == null) return false;
        XElement? properties = source.Styles.Find(OdfStyleFamily.Graphic, styleName)?.Element
            .Element(OdfNamespaces.Style + "graphic-properties");
        if (properties == null) return false;
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        string? stroke = (string?)properties.Attribute(OdfNamespaces.Draw + "stroke");
        if (fill != null && fill != "solid" || stroke != null && stroke != "solid") return true;
        if (fill == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out _)) return true;
        if (stroke == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Svg + "stroke-color"), out _)) return true;
        return properties.Attributes().Any(attribute =>
            (attribute.Name.Namespace == OdfNamespaces.Draw || attribute.Name.Namespace == OdfNamespaces.Svg) &&
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color" &&
            attribute.Name != OdfNamespaces.Draw + "stroke" && attribute.Name != OdfNamespaces.Svg + "stroke-color" &&
            attribute.Name != OdfNamespaces.Svg + "stroke-width");
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

    private static (bool Override, bool Loss, OdfColor? Color) ReadOdpSlideBackground(
        OdpPresentation source, OdpSlide slide) {
        string? styleName = (string?)slide.Element.Attribute(OdfNamespaces.Draw + "style-name");
        if (styleName == null) return default;
        XElement? properties = source.Styles.Find(OdfStyleFamily.DrawingPage, styleName)?.Element
            .Element(OdfNamespaces.Style + "drawing-page-properties");
        if (properties == null) return default;
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        if (fill == null) return default;
        OdfColor? color = fill == "solid" &&
            OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out OdfColor parsed)
            ? parsed : (OdfColor?)null;
        bool loss = fill != "none" && !color.HasValue || properties.Attributes().Any(attribute =>
            attribute.Name.Namespace == OdfNamespaces.Draw &&
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color");
        return (true, loss, color);
    }
}
