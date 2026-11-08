using System.Xml.Linq;
using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static readonly XNamespace Sf = IWorkLegacyXmlSource.Sf;
    private static readonly XNamespace Sfa = IWorkLegacyXmlSource.Sfa;

    internal static IWorkPagesProjection ReadPages(IWorkSourceDocument source, IWorkLegacyXmlSource legacy) {
        XElement? body = legacy.Root.Element(Sf + "text-storage")?.Element(Sf + "text-body");
        var budget = new IWorkProjectionBudget(source.Options);
        var catalog = new IWorkLegacyXmlCatalog(source, legacy, budget);
        if (legacy.Root.Element(Sf + "text-storage")?.Element(Sf + "footnotes") != null)
            catalog.AddLoss("IWORK_LEGACY_FOOTNOTES_OMITTED",
                "Pages '09 footnotes and their inline fields are outside the editable projection.");
        var tables = new List<IWorkTable>();
        var images = new List<IWorkImageAsset>();
        var boxes = new List<IWorkTextBox>();
        var drawables = new List<IWorkPagesDrawable>();
        int materializedCells = 0;
        var attachments = new HashSet<XElement>();
        if (body != null) {
            foreach (XElement reference in body.Descendants(Sf + "attachment-ref")) {
                source.CancellationToken.ThrowIfCancellationRequested();
                budget.AddDrawableReferences(1);
                XElement? attachment = catalog.Resolve(reference, "attachment");
                XElement? drawable = attachment?.Elements().FirstOrDefault(element =>
                    element.Name == Sf + "tabular-info" || element.Name == Sf + "media");
                if (attachment == null || drawable == null || !attachments.Add(attachment)) {
                    LegacyDrawableLoss(catalog);
                    continue;
                }
                if (AddDrawable(drawable, null)) catalog.RegisterInlineObject(reference, attachment, drawable);
            }
        }
        IWorkTextContent content = catalog.ReadText(body);
        IReadOnlyList<IWorkPagesSection> sections = ReadLegacyPagesRunningStories(legacy, body, catalog);
        XNamespace pages = IWorkLegacyXmlSource.Pages;
        foreach (XElement page in legacy.Root.Element(pages + "drawables")?.Elements(pages + "page-group")
            ?? Enumerable.Empty<XElement>()) {
            source.CancellationToken.ThrowIfCancellationRequested();
            int? pageIndex = int.TryParse((string?)page.Attribute(pages + "page"), NumberStyles.None,
                CultureInfo.InvariantCulture, out int number) && number > 0 ? number : null;
            foreach (XElement drawable in page.Elements()) {
                source.CancellationToken.ThrowIfCancellationRequested();
                budget.AddDrawableReferences(1);
                AddDrawable(drawable, pageIndex);
            }
        }
        catalog.AddLoss("IWORK_LEGACY_PAGES_CONTENT_LIMITED",
            "The iWork '09 Pages projection recovers selected text, formatting, tables, objects, page dimensions, "
            + "and qualified single-section default headers and footers. Other section layouts, running-story "
            + "selection, and unqualified objects remain outside the editable projection.");
        return new IWorkPagesProjection(source, content, sections,
            boxes, images, tables, drawables, ReadLegacyPageLayout(legacy), catalog.Diagnostics,
            supportsEditableReconstruction: false);

        bool AddDrawable(XElement drawable, int? pageIndex) {
            if (drawable.Name == Sf + "tabular-info") {
                budget.AddTable();
                IWorkTable? table = ReadLegacyTable(drawable, source, budget, catalog, ref materializedCells);
                if (table != null) {
                    tables.Add(table);
                    drawables.Add(new IWorkPagesDrawable(table, pageIndex));
                    return true;
                }
            } else if (drawable.Name == Sf + "media") {
                IWorkImageAsset? image = ReadLegacyImage(source, drawable, budget, catalog);
                if (image != null) {
                    images.Add(image);
                    drawables.Add(new IWorkPagesDrawable(image, pageIndex));
                    return true;
                }
            } else if (drawable.Name == Sf + "drawable-shape" || drawable.Name == Sf + "shape") {
                IWorkTextBox? box = ReadLegacyTextBox(drawable, catalog);
                if (box != null) {
                    boxes.Add(box);
                    drawables.Add(new IWorkPagesDrawable(box, pageIndex));
                    return true;
                }
            }
            LegacyDrawableLoss(catalog);
            return false;
        }
    }

    private static IWorkPageLayout? ReadLegacyPageLayout(IWorkLegacyXmlSource legacy) {
        XNamespace pages = IWorkLegacyXmlSource.Pages;
        XElement? print = legacy.Root.Element(pages + "slprint-info");
        XElement? margins = print?.Element(Sf + "page-margins");
        double? width = IWorkLegacyXmlCatalog.AttributeNumber(print, pages + "page-width");
        double? height = IWorkLegacyXmlCatalog.AttributeNumber(print, pages + "page-height");
        double? left = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "left");
        double? right = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "right");
        double? top = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "top");
        double? bottom = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "bottom");
        double? header = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "header");
        double? footer = IWorkLegacyXmlCatalog.AttributeNumber(margins, Sf + "footer");
        if (width is not > 0 || height is not > 0 || left is not >= 0 || right is not >= 0
            || top is not >= 0 || bottom is not >= 0 || header is not >= 0 || footer is not >= 0
            || left + right >= width || top + bottom >= height) return null;
        return new IWorkPageLayout(width.Value, height.Value, left.Value, right.Value, top.Value, bottom.Value,
            header.Value, footer.Value, width > height);
    }

    private static IWorkDiagnostic LegacyLoss(IWorkLegacyXmlSource legacy, string code, string message) =>
        new(IWorkDiagnosticSeverity.Warning, code, message, legacy.EntryPath);
}
