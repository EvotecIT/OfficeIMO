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

    private static readonly Lazy<string?> DefaultPowerPointNotesMasterXml = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return baseline.OpenXmlDocument.PresentationPart?.NotesMasterPart?.NotesMaster?.OuterXml;
    });

    private static readonly Lazy<string?> DefaultPowerPointHandoutMasterXml = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return baseline.OpenXmlDocument.PresentationPart?.HandoutMasterPart?.HandoutMaster?.OuterXml;
    });

    private static readonly Lazy<string?> DefaultPowerPointShowPropertiesXml = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return baseline.OpenXmlDocument.PresentationPart?.PresentationPropertiesPart?
            .PresentationProperties?.ShowProperties?.OuterXml;
    });

    private static readonly Lazy<string?> DefaultPowerPointViewPropertiesXml = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return baseline.OpenXmlDocument.PresentationPart?.ViewPropertiesPart?.ViewProperties?.OuterXml;
    });

    private static readonly Lazy<HashSet<string>> DefaultPowerPointMasterColorMaps = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return new HashSet<string>(baseline.OpenXmlDocument.PresentationPart!.SlideMasterParts
            .Select(master => master.SlideMaster?.ColorMap?.OuterXml ?? string.Empty), StringComparer.Ordinal);
    });

    private static readonly Lazy<HashSet<string>> DefaultPowerPointMasterTextStyles = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return new HashSet<string>(baseline.OpenXmlDocument.PresentationPart!.SlideMasterParts
            .Select(master => master.SlideMaster?.TextStyles?.OuterXml ?? string.Empty), StringComparer.Ordinal);
    });

    private static readonly Lazy<HashSet<string>> DefaultPowerPointMasterMetadata = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return new HashSet<string>(baseline.OpenXmlDocument.PresentationPart!.SlideMasterParts
            .Select(master => MasterMetadataSignature(master.SlideMaster)), StringComparer.Ordinal);
    });

    private static readonly Lazy<HashSet<string>> DefaultPowerPointThemes = new(() => {
        using PowerPointPresentation baseline = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PresentationPart? presentation = baseline.OpenXmlDocument.PresentationPart;
        return new HashSet<string>(new[] { presentation?.ThemePart?.Theme?.OuterXml }
            .Concat(presentation?.SlideMasterParts.Select(master => master.ThemePart?.Theme?.OuterXml)
                ?? Enumerable.Empty<string?>())
            .Concat(new[] { presentation?.NotesMasterPart?.ThemePart?.Theme?.OuterXml,
                presentation?.HandoutMasterPart?.ThemePart?.Theme?.OuterXml })
            .Where(xml => xml != null).Select(xml => xml!), StringComparer.Ordinal);
    });

    private static int CountUnmappedPowerPointThemes(PresentationPart? presentation) {
        if (presentation == null) return 0;
        int changedThemes = new[] { presentation.ThemePart?.Theme?.OuterXml,
                presentation.NotesMasterPart?.ThemePart?.Theme?.OuterXml,
                presentation.HandoutMasterPart?.ThemePart?.Theme?.OuterXml }
            .Concat(presentation.SlideMasterParts.Select(master => master.ThemePart?.Theme?.OuterXml))
            .Count(xml => xml != null && !DefaultPowerPointThemes.Value.Contains(xml));
        int overrides = presentation.SlideParts.Count(slide =>
                slide.ThemeOverridePart?.ThemeOverride != null ||
                HasAuthoredColorMapOverride(slide.Slide?.ColorMapOverride) ||
                slide.NotesSlidePart?.ThemeOverridePart?.ThemeOverride != null ||
                HasAuthoredColorMapOverride(slide.NotesSlidePart?.NotesSlide?.ColorMapOverride)) +
            presentation.SlideMasterParts.Sum(master => master.SlideLayoutParts.Count(layout =>
                layout.ThemeOverridePart?.ThemeOverride != null ||
                HasAuthoredColorMapOverride(layout.SlideLayout?.ColorMapOverride)));
        return changedThemes + overrides;
    }

    private static bool HasAuthoredColorMapOverride(P.ColorMapOverride? colorMap) =>
        colorMap != null && (colorMap.HasAttributes || colorMap.ChildElements.Any(child =>
            child is not A.MasterColorMapping));

    private static int CountUnmappedPowerPointEmbeddedFonts(PresentationPart? presentation) =>
        presentation?.Presentation?.Descendants()
            .Count(element => element.LocalName == "embeddedFont") ?? 0;

    private static string MasterMetadataSignature(P.SlideMaster? master) => string.Join(";",
        (master?.GetAttributes() ?? new List<DocumentFormat.OpenXml.OpenXmlAttribute>())
            .Concat(master?.CommonSlideData?.GetAttributes() ?? new List<DocumentFormat.OpenXml.OpenXmlAttribute>())
            .Select(attribute => attribute.NamespaceUri + "|" + attribute.LocalName + "|" + attribute.Value)
            .OrderBy(value => value, StringComparer.Ordinal));

    private static readonly Lazy<(string Master, string Layout)> DefaultOdpMasterLayoutNames = new(() => {
        OdpPresentation baseline = OdpPresentation.Create();
        baseline.AddSlide("Baseline");
        return (baseline.MasterPages[0].Name, baseline.Layouts[0].Name);
    });

    private static int CountUnmappedPowerPointNotesMaster(PresentationPart? presentation) {
        string? sourceXml = presentation?.NotesMasterPart?.NotesMaster?.OuterXml;
        if (sourceXml == null) return 0;
        string? defaultXml = DefaultPowerPointNotesMasterXml.Value;
        return defaultXml != null && string.Equals(sourceXml, defaultXml, StringComparison.Ordinal)
            ? 0 : 1;
    }

    private static int CountUnmappedPowerPointHandoutMaster(PresentationPart? presentation) {
        string? sourceXml = presentation?.HandoutMasterPart?.HandoutMaster?.OuterXml;
        if (sourceXml == null) return 0;
        return string.Equals(sourceXml, DefaultPowerPointHandoutMasterXml.Value, StringComparison.Ordinal)
            ? 0 : 1;
    }

    private static int CountUnmappedPowerPointShowProperties(PresentationPart? presentation) {
        string? sourceXml = presentation?.PresentationPropertiesPart?.PresentationProperties?
            .ShowProperties?.OuterXml;
        if (sourceXml == null) return 0;
        return string.Equals(sourceXml, DefaultPowerPointShowPropertiesXml.Value, StringComparison.Ordinal)
            ? 0 : 1;
    }

    private static int CountUnmappedPowerPointViewProperties(PresentationPart? presentation) {
        string? sourceXml = presentation?.ViewPropertiesPart?.ViewProperties?.OuterXml;
        if (sourceXml == null) return 0;
        return string.Equals(sourceXml, DefaultPowerPointViewPropertiesXml.Value, StringComparison.Ordinal)
            ? 0 : 1;
    }

    private static bool MapPowerPointMasterAndLayout(PresentationPart? presentation, SlideId? slideId,
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
        if (HasUnmappedMasterOrLayoutContent(masterPart, layoutPart)) approximatedMasterLayouts++;

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

    private static bool HasUnmappedMasterOrLayoutContent(SlideMasterPart masterPart, SlideLayoutPart layoutPart) =>
        HasDrawingContent(masterPart.SlideMaster?.CommonSlideData?.ShapeTree) ||
        HasDrawingContent(layoutPart.SlideLayout?.CommonSlideData?.ShapeTree) ||
        HasAuthoredPowerPointLayoutContent(layoutPart) ||
        HasAuthoredPowerPointMasterColorMap(masterPart) ||
        HasAuthoredPowerPointMasterTextStyles(masterPart) ||
        HasAuthoredPowerPointMasterMetadata(masterPart) ||
        HasAuthoredPowerPointMasterBehavior(masterPart) ||
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

    private static bool HasAuthoredPowerPointMasterColorMap(SlideMasterPart masterPart) =>
        !DefaultPowerPointMasterColorMaps.Value.Contains(masterPart.SlideMaster?.ColorMap?.OuterXml ?? string.Empty);

    private static bool HasAuthoredPowerPointMasterTextStyles(SlideMasterPart masterPart) =>
        !DefaultPowerPointMasterTextStyles.Value.Contains(masterPart.SlideMaster?.TextStyles?.OuterXml ?? string.Empty);

    private static bool HasAuthoredPowerPointMasterMetadata(SlideMasterPart masterPart) =>
        !DefaultPowerPointMasterMetadata.Value.Contains(MasterMetadataSignature(masterPart.SlideMaster));

    private static bool HasAuthoredPowerPointMasterBehavior(SlideMasterPart masterPart) =>
        masterPart.SlideMaster?.ChildElements.Any(child =>
            child.LocalName is "transition" or "timing" or "hf") == true;

    private static int CountUnmappedUnusedPowerPointMastersAndLayouts(PresentationPart? presentation,
        IReadOnlyDictionary<SlideMasterPart, string> usedMasters,
        IReadOnlyDictionary<SlideLayoutPart, string> usedLayouts) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (SlideMasterPart master in presentation.SlideMasterParts) {
            if (!usedMasters.ContainsKey(master) &&
                (HasDrawingContent(master.SlideMaster?.CommonSlideData?.ShapeTree) ||
                 HasAuthoredPowerPointMasterColorMap(master) ||
                 HasAuthoredPowerPointMasterMetadata(master) ||
                 HasAuthoredPowerPointMasterBehavior(master) ||
                 master.SlideMaster?.CommonSlideData?.Background != null ||
                 HasAuthoredPowerPointMasterTextStyles(master)))
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
        bool hasUnsupportedAttributes = styles.Descendants(style + "master-page").Any(master =>
            master.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                attribute.Name != style + "name" && attribute.Name != style + "page-layout-name" &&
                attribute.Name != draw + "style-name")) ||
            styles.Descendants(style + "presentation-page-layout").Any(layout =>
                layout.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    attribute.Name != style + "name"));
        XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        XElement[] pageLayouts = styles.Descendants(style + "page-layout").ToArray();
        bool hasUnsupportedPageLayout = pageLayouts.Length > 1 || pageLayouts.Any(layout => {
            XElement? properties = layout.Element(style + "page-layout-properties");
            return layout.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    attribute.Name != style + "name") ||
                properties == null || properties.HasElements ||
                properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    attribute.Name != fo + "page-width" && attribute.Name != fo + "page-height" &&
                    !(attribute.Name == style + "print-orientation" && attribute.Value == "landscape") &&
                    !(attribute.Name == fo + "margin" && IsZeroOdfPageMargin(attribute.Value)));
        });
        bool hasUnknownPageLayoutReference = styles.Descendants(style + "master-page").Any(master => {
            string? name = (string?)master.Attribute(style + "page-layout-name");
            return !string.IsNullOrWhiteSpace(name) && !pageLayouts.Any(layout =>
                string.Equals((string?)layout.Attribute(style + "name"), name, StringComparison.Ordinal));
        });
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
        bool hasAuthoredNames = masters.Count == 1 &&
            !string.Equals(masters[0].Name, DefaultOdpMasterLayoutNames.Value.Master, StringComparison.Ordinal) ||
            layouts.Count == 1 &&
            !string.Equals(layouts[0].Name, DefaultOdpMasterLayoutNames.Value.Layout, StringComparison.Ordinal);
        return hasContent || hasUnsupportedAttributes || hasUnsupportedPageLayout ||
            hasUnknownPageLayoutReference || hasUnsupportedBackground ||
            hasAuthoredNames || masters.Count > 1 || layouts.Count > 1
            ? masters.Count + layouts.Count
            : 0;
    }

    private static bool IsZeroOdfPageMargin(string value) =>
        !string.IsNullOrWhiteSpace(value) && OdfLength.Parse(value).TryToPoints(out double points) &&
        points == 0D;
}
