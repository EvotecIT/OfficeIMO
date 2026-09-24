using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static readonly Lazy<HashSet<string>> DefaultPowerPointLayoutXml = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return new HashSet<string>(baseline.OpenXmlDocument.PresentationPart!.SlideMasterParts
            .SelectMany(master => master.SlideLayoutParts)
            .Select(layout => layout.SlideLayout?.OuterXml)
            .Where(xml => xml != null).Select(xml => xml!), StringComparer.Ordinal);
    });

    private static bool MapPowerPointMasterAndLayout(PresentationPart? presentation, SlideId? slideId,
        bool hasSlideText,
        OdpPresentation target, OdpSlide targetSlide,
        Dictionary<SlideMasterPart, string> masterNames, Dictionary<SlideLayoutPart, string> layoutNames,
        HashSet<SlideMasterPart> solidMasters,
        ref int mappedMasters, ref int mappedLayouts, ref int mappedMasterBackgrounds,
        ref int approximatedMasterLayouts, out bool suppressInheritedBackground) {
        suppressInheritedBackground = false;
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

        bool slideOverride = slidePart.Slide?.CommonSlideData?.Background != null;
        bool layoutOverride = layoutPart.SlideLayout?.CommonSlideData?.Background != null;
        suppressInheritedBackground = solidMasters.Contains(masterPart) && (slideOverride || layoutOverride);
        return solidMasters.Contains(masterPart) && !slideOverride && !layoutOverride;
    }

    private static bool HasUnmappedMasterOrLayoutContent(SlideMasterPart masterPart, SlideLayoutPart layoutPart,
        bool hasSlideText) =>
        HasDrawingContent(masterPart.SlideMaster?.CommonSlideData?.ShapeTree) ||
        HasDrawingContent(layoutPart.SlideLayout?.CommonSlideData?.ShapeTree) ||
        HasAuthoredPowerPointLayoutContent(layoutPart) ||
        hasSlideText && masterPart.SlideMaster?.TextStyles?.ChildElements.Any(style =>
            style.HasAttributes || style.HasChildren) == true ||
        masterPart.SlideMaster?.CommonSlideData?.Background != null &&
            !TryGetDirectMasterBackground(masterPart, out _) ||
        layoutPart.SlideLayout?.CommonSlideData?.Background != null;

    private static bool TryGetDirectMasterBackground(SlideMasterPart masterPart, out OdfColor color) {
        P.BackgroundProperties? properties = masterPart.SlideMaster?.CommonSlideData?.Background?.BackgroundProperties;
        A.SolidFill? solid = properties?.GetFirstChild<A.SolidFill>();
        A.RgbColorModelHex? rgb = solid?.RgbColorModelHex;
        if (properties != null && !properties.HasAttributes && properties.ChildElements.Count == 1 &&
            solid?.ChildElements.Count == 1 && rgb?.ChildElements.Count == 0 &&
            OdfColor.TryParse(rgb.Val?.Value, out color)) return true;
        color = default;
        return false;
    }

    private static bool HasDrawingContent(P.ShapeTree? tree) => tree?.ChildElements.Any(child =>
        child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties) == true;

    private static int CountUnmappedUnusedPowerPointMastersAndLayouts(PresentationPart? presentation,
        IReadOnlyDictionary<SlideMasterPart, string> usedMasters,
        IReadOnlyDictionary<SlideLayoutPart, string> usedLayouts) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (SlideMasterPart master in presentation.SlideMasterParts) {
            if (!usedMasters.ContainsKey(master) &&
                (HasDrawingContent(master.SlideMaster?.CommonSlideData?.ShapeTree) ||
                 master.SlideMaster?.CommonSlideData?.Background != null ||
                 master.SlideMaster?.TextStyles?.ChildElements.Any(style => style.HasAttributes || style.HasChildren) == true))
                count++;
            foreach (SlideLayoutPart layout in master.SlideLayoutParts) {
                if (!usedLayouts.ContainsKey(layout) &&
                    (HasAuthoredPowerPointLayoutContent(layout) ||
                     layout.SlideLayout?.CommonSlideData?.Background != null)) count++;
            }
        }
        return count;
    }

    private static bool HasAuthoredPowerPointLayoutContent(SlideLayoutPart layout) {
        P.SlideLayout? source = layout.SlideLayout;
        if (source == null) return false;
        // The stock layouts are unedited skeletons. Any change to an unused
        // layout's placeholder geometry, appearance, or metadata is lost.
        return !DefaultPowerPointLayoutXml.Value.Contains(source.OuterXml);
    }

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
            XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage,
                styleName, style + "drawing-page-properties", "styles.xml");
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
