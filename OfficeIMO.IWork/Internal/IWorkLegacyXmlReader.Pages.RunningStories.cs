using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    /// <summary>Reads the qualified single-section default stories without selecting inactive page templates.</summary>
    private static IReadOnlyList<IWorkPagesSection> ReadLegacyPagesRunningStories(
        IWorkLegacyXmlSource legacy, XElement? body, IWorkLegacyXmlCatalog catalog) {
        XElement[] headers = legacy.Root.Elements(Sf + "headers").Take(2).ToArray();
        XElement[] footers = legacy.Root.Elements(Sf + "footers").Take(2).ToArray();
        if (headers.Length == 0 && footers.Length == 0) return Array.Empty<IWorkPagesSection>();

        XElement[] sections = body?.Elements(Sf + "section").Take(2).ToArray() ?? Array.Empty<XElement>();
        if (sections.Length != 1 || sections[0].Descendants(Sf + "section").Any()
            || body!.Elements().Any(element => element.Name != Sf + "section"
                && element.Name != Sf + "page-start" && element.Name != Sf + "container-hint"
                && element.Name != Sf + "insertion-point")) {
            LegacyRunningStoryLoss(catalog);
            return Array.Empty<IWorkPagesSection>();
        }

        XElement? style = catalog.Style(sections[0], (string?)sections[0].Attribute(Sf + "style"), "sectionstyle");
        XElement? master = catalog.Property(style, "oddPageMaster");
        XElement[] masters = master?.Elements().Take(2).ToArray() ?? Array.Empty<XElement>();
        if (style == null
            || LegacyRunningStoryFlag(catalog.Property(style, "pageMasterFirstPageDifferent")) != false
            || LegacyRunningStoryFlag(catalog.Property(style, "pageMasterLeftRightPagesDifferent")) != false
            || masters.Length != 1 || masters[0].Name != Sf + "pagemaster"
            || masters[0].HasAttributes || masters[0].HasElements) {
            LegacyRunningStoryLoss(catalog);
            return Array.Empty<IWorkPagesSection>();
        }

        XNamespace pages = IWorkLegacyXmlSource.Pages;
        XElement? publication = legacy.Root.Element(pages + "publication-info");
        IReadOnlyList<IWorkTextContent>? defaultHeaders = ReadDefaultStory(headers, "header",
            "SFWPDefaultOddHeaderIdentifier", "kSFWPHasHeadersProperty");
        IReadOnlyList<IWorkTextContent>? defaultFooters = ReadDefaultStory(footers, "footer",
            "SFWPDefaultOddFooterIdentifier", "kSFWPHasFootersProperty");
        return defaultHeaders == null && defaultFooters == null
            ? Array.Empty<IWorkPagesSection>()
            : new[] { new IWorkPagesSection(0, null, null, null, null, defaultHeaders, defaultFooters) };

        IReadOnlyList<IWorkTextContent>? ReadDefaultStory(XElement[] containers, string kind,
            string identifier, string enabledProperty) {
            bool? enabled = LegacyRunningStoryFlag(publication?.Element(pages + enabledProperty));
            if (enabled == false) return null;
            XElement[] stories = containers.Length == 1
                ? containers[0].Elements(Sf + kind)
                    .Where(element => (string?)element.Attribute(Sf + "name") == identifier).Take(2).ToArray()
                : Array.Empty<XElement>();
            XElement[] storages = stories.Length == 1
                ? stories[0].Elements(Sf + "text-storage").Take(2).ToArray() : Array.Empty<XElement>();
            XElement[] bodies = storages.Length == 1
                ? storages[0].Elements(Sf + "text-body").Take(2).ToArray() : Array.Empty<XElement>();
            if (enabled != true || bodies.Length != 1) {
                LegacyRunningStoryLoss(catalog);
                return null;
            }
            return new[] { catalog.ReadText(bodies[0]) };
        }
    }

    private static bool? LegacyRunningStoryFlag(XElement? property) {
        XElement[] values = property?.Elements().Take(2).ToArray() ?? Array.Empty<XElement>();
        if (values.Length != 1 || (values[0].Name != Sf + "number"
            && values[0].Name != IWorkLegacyXmlSource.Pages + "number")) return null;
        return (string?)values[0].Attribute(Sfa + "number") switch {
            "0" => false,
            "1" => true,
            _ => null
        };
    }

    private static void LegacyRunningStoryLoss(IWorkLegacyXmlCatalog catalog) =>
        catalog.AddLoss("IWORK_LEGACY_PAGES_HEADER_FOOTER_OMITTED",
            "A Pages '09 header or footer has unqualified section selection, enablement, or default storage; "
            + "only qualified single-section default running stories are reconstructed.");
}
