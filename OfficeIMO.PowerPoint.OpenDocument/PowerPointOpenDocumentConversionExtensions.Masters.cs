using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static bool MapPowerPointMasterAndLayout(PresentationPart? presentation, SlideId? slideId,
        bool hasSlideText,
        OdpPresentation target, OdpSlide targetSlide,
        Dictionary<SlideMasterPart, string> masterNames, Dictionary<SlideLayoutPart, string> layoutNames,
        HashSet<SlideMasterPart> solidMasters,
        ref int mappedMasters, ref int mappedLayouts, ref int mappedMasterBackgrounds,
        ref int approximatedMasterLayouts) {
        string? relationshipId = slideId?.RelationshipId?.Value;
        if (presentation == null || relationshipId == null || relationshipId.Length == 0 ||
            presentation.GetPartById(relationshipId) is not SlidePart slidePart) return false;
        SlideLayoutPart? layoutPart = slidePart.SlideLayoutPart;
        SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
        if (masterPart == null || layoutPart == null) return false;
        if (HasUnmappedMasterOrLayoutContent(masterPart, layoutPart, hasSlideText)) approximatedMasterLayouts++;

        if (!masterNames.TryGetValue(masterPart, out string? masterName)) {
            OdpMasterPage targetMaster = masterNames.Count == 0
                ? target.MasterPages[0]
                : target.AddMasterPage("Master" + (masterNames.Count + 1));
            masterName = targetMaster.Name;
            masterNames.Add(masterPart, masterName);
            if (TryGetDirectMasterBackground(masterPart, out OdfColor color)) {
                targetMaster.BackgroundColor = color;
                solidMasters.Add(masterPart);
                mappedMasterBackgrounds++;
            }
        }
        targetSlide.MasterPageName = masterName;
        mappedMasters++;

        if (!layoutNames.TryGetValue(layoutPart, out string? layoutName)) {
            layoutName = layoutNames.Count == 0
                ? target.Layouts[0].Name
                : target.AddLayout("Layout" + (layoutNames.Count + 1)).Name;
            layoutNames.Add(layoutPart, layoutName);
        }
        targetSlide.LayoutName = layoutName;
        mappedLayouts++;

        return solidMasters.Contains(masterPart) &&
            slidePart.Slide?.CommonSlideData?.Background == null &&
            layoutPart.SlideLayout?.CommonSlideData?.Background == null;
    }

    private static bool HasUnmappedMasterOrLayoutContent(SlideMasterPart masterPart, SlideLayoutPart layoutPart,
        bool hasSlideText) =>
        HasDrawingContent(masterPart.SlideMaster?.CommonSlideData?.ShapeTree) ||
        HasDrawingContent(layoutPart.SlideLayout?.CommonSlideData?.ShapeTree) ||
        hasSlideText && masterPart.SlideMaster?.TextStyles?.ChildElements.Count > 0 ||
        masterPart.SlideMaster?.CommonSlideData?.Background != null &&
            !TryGetDirectMasterBackground(masterPart, out _) ||
        layoutPart.SlideLayout?.CommonSlideData?.Background != null;

    private static bool TryGetDirectMasterBackground(SlideMasterPart masterPart, out OdfColor color) {
        A.SolidFill? solid = masterPart.SlideMaster?.CommonSlideData?.Background?
            .BackgroundProperties?.GetFirstChild<A.SolidFill>();
        A.RgbColorModelHex? rgb = solid?.RgbColorModelHex;
        if (rgb?.ChildElements.Count == 0 && OdfColor.TryParse(rgb.Val?.Value, out color)) return true;
        color = default;
        return false;
    }

    private static bool HasDrawingContent(P.ShapeTree? tree) => tree?.ChildElements.Any(child =>
        child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties) == true;

    private static int CountUnmappedOdpMasterLayouts(OdpPresentation source) {
        IReadOnlyList<OdpMasterPage> masters = source.MasterPages;
        IReadOnlyList<OdpPresentationLayout> layouts = source.Layouts;
        if (masters.Count == 0 && layouts.Count == 0) return 0;
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XDocument styles = source.Package.GetXml("styles.xml");
        bool hasContent = styles.Descendants(style + "master-page").Any(master => master.Elements().Any()) ||
            styles.Descendants(style + "presentation-page-layout").Any(layout =>
                layout.Elements(presentation + "placeholder").Any() || layout.Elements().Any(element =>
                    element.Name != presentation + "placeholder"));
        bool hasUnsupportedBackground = styles.Descendants(style + "master-page").Any(master => {
            string? styleName = (string?)master.Attribute(draw + "style-name");
            if (styleName == null) return false;
            XElement? drawingStyle = styles.Descendants(style + "style").FirstOrDefault(candidate =>
                (string?)candidate.Attribute(style + "name") == styleName &&
                (string?)candidate.Attribute(style + "family") == "drawing-page");
            XElement? properties = drawingStyle?.Element(style + "drawing-page-properties");
            if (properties == null) return false;
            string? fill = (string?)properties.Attribute(draw + "fill");
            if (fill != null && fill != "none" && fill != "solid") return true;
            if (fill == "solid" && !OdfColor.TryParse((string?)properties.Attribute(draw + "fill-color"), out _))
                return true;
            return properties.Attributes().Any(attribute => attribute.Name != draw + "fill" &&
                attribute.Name != draw + "fill-color");
        });
        return hasContent || hasUnsupportedBackground || masters.Count > 1 || layouts.Count > 1
            ? masters.Count + layouts.Count
            : 0;
    }
}
