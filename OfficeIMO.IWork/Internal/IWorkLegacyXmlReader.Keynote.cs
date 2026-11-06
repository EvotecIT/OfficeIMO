using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static readonly XNamespace Key = IWorkLegacyXmlSource.Keynote;
    private static readonly string[] LegacyUnsupportedTransforms =
        { "shearXAngle", "shearYAngle", "horizontalFlip", "verticalFlip" };

    internal static IWorkKeynoteProjection ReadKeynote(IWorkSourceDocument source, IWorkLegacyXmlSource legacy) {
        XElement? slideList = legacy.Root.Element(Key + "slide-list");
        var budget = new IWorkProjectionBudget(source.Options);
        var catalog = new IWorkLegacyXmlCatalog(source, legacy, budget);
        var slides = new List<IWorkKeynoteSlide>();
        int materializedCells = 0;
        if (slideList != null) {
            foreach (XElement slide in slideList.Elements(Key + "slide")) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (slides.Count >= source.Options.MaxProjectedSlides)
                    throw new InvalidDataException("The Keynote '09 presentation exceeds the configured slide limit.");
                slides.Add(ReadKeynoteSlide(source, slide, slides.Count + 1, budget, catalog, ref materializedCells));
            }
        }

        XElement? size = legacy.Root.Element(Key + "size");
        double? width = KeynoteNumber(size, "w");
        double? height = KeynoteNumber(size, "h");
        IWorkCanvasSize? slideSize = width is > 0 && height is > 0 ? new(width.Value, height.Value) : null;
        var diagnostics = new[] {
            LegacyLoss(legacy, "IWORK_LEGACY_KEYNOTE_CONTENT_LIMITED",
                "The Keynote '09 profile recovers ordered slides, selected text and formatting, "
                + "and presenter notes. Masters, unqualified drawables, and playback settings remain outside the editable projection.")
        };
        return new IWorkKeynoteProjection(source, slides, slideSize, diagnostics.Concat(catalog.Diagnostics).ToArray(),
            supportsEditableReconstruction: false);
    }

    private static IWorkKeynoteSlide ReadKeynoteSlide(IWorkSourceDocument source, XElement slide,
        int index, IWorkProjectionBudget budget, IWorkLegacyXmlCatalog catalog, ref int materializedCells) {
        string name = (string?)slide.Attribute(Key + "name") ?? string.Empty;
        budget.AddTextCharacters(name.Length);
        IWorkTextBox? title = null;
        var boxes = new List<IWorkTextBox>();
        var images = new List<IWorkImageAsset>();
        var tables = new List<IWorkTable>();
        var drawables = new List<IWorkKeynoteDrawable>();
        XElement? page = slide.Element(Key + "page");
        XElement? titlePlaceholder = slide.Element(Key + "title-placeholder");
        XElement? bodyPlaceholder = slide.Element(Key + "body-placeholder");
        XName idAttribute = Sfa + "ID";
        XName referenceAttribute = Sfa + "IDREF";
        string? titleId = (string?)titlePlaceholder?.Attribute(idAttribute);
        string? bodyId = (string?)bodyPlaceholder?.Attribute(idAttribute);
        var acceptedPlaceholders = new HashSet<XElement>();
        foreach (XElement drawable in LegacyCanvasDrawables(page)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            budget.AddDrawableReferences(1);
            bool isTitle = drawable.Name == Sf + "title-placeholder-ref";
            if (isTitle || drawable.Name == Sf + "body-placeholder-ref") {
                XElement? placeholder = isTitle ? titlePlaceholder : bodyPlaceholder;
                string? expectedId = isTitle ? titleId : bodyId;
                if (placeholder == null || string.IsNullOrEmpty(expectedId)
                    || (string?)drawable.Attribute(referenceAttribute) != expectedId
                    || !acceptedPlaceholders.Add(placeholder)) continue;
                IWorkTextBox? box = ReadLegacyTextBox(placeholder, catalog);
                if (box?.Content.Paragraphs.Any(paragraph => paragraph.Runs.Count > 0) == true) {
                    if (isTitle) title = box;
                    else boxes.Add(box);
                    drawables.Add(new IWorkKeynoteDrawable(box, isTitle));
                }
            } else if (drawable.Name == Sf + "shape" || drawable.Name == Sf + "drawable-shape") {
                IWorkTextBox? box = ReadLegacyTextBox(drawable, catalog);
                if (box != null) {
                    boxes.Add(box);
                    drawables.Add(new IWorkKeynoteDrawable(box, false));
                } else LegacyDrawableLoss(catalog);
            } else if (drawable.Name == Sf + "tabular-info") {
                budget.AddTable();
                IWorkTable? table = ReadLegacyTable(drawable, source, budget, catalog, ref materializedCells);
                if (table != null) {
                    tables.Add(table);
                    drawables.Add(new IWorkKeynoteDrawable(table));
                } else LegacyDrawableLoss(catalog);
            } else if (drawable.Name == Sf + "media") {
                IWorkImageAsset? image = ReadLegacyImage(source, drawable, budget, catalog);
                if (image != null) {
                    images.Add(image);
                    drawables.Add(new IWorkKeynoteDrawable(image));
                }
            } else {
                LegacyDrawableLoss(catalog);
            }
        }

        XElement? notesBody = slide.Element(Key + "notes")
            ?.Element(Sf + "text-storage")?.Element(Sf + "text-body");
        IWorkTextContent notes = catalog.ReadText(notesBody);
        return new IWorkKeynoteSlide(index, name, title, boxes, notes, images,
            tables, drawables, isSkipped: false);
    }

    internal static IWorkGeometry? ReadLegacyGeometry(XElement placeholder, IWorkLegacyXmlCatalog? catalog = null) {
        XElement? geometry = placeholder.Element(Sf + "geometry");
        XElement? size = geometry?.Element(Sf + "size");
        XElement? position = geometry?.Element(Sf + "position");
        double? width = KeynoteNumber(size, "w");
        double? height = KeynoteNumber(size, "h");
        double? x = KeynoteNumber(position, "x");
        double? y = KeynoteNumber(position, "y");
        double? angle = IWorkLegacyXmlCatalog.AttributeNumber(geometry, Sf + "angle");
        bool unsupportedTransform = LegacyUnsupportedTransforms.Any(name => geometry?.Attribute(Sf + name) is XAttribute attribute
                && attribute.Value is not ("0" or "false"));
        if (unsupportedTransform || geometry?.Attribute(Sf + "angle") != null && !angle.HasValue) {
            catalog?.AddLoss("IWORK_LEGACY_GEOMETRY_UNSUPPORTED",
                "A selected XML drawable uses an unqualified transform; its geometry is omitted.");
            return null;
        }
        return width is > 0 && height is > 0 && x.HasValue && y.HasValue
            ? new IWorkGeometry(x.Value, y.Value, width.Value, height.Value, -(angle ?? 0))
            : null;
    }

    private static double? KeynoteNumber(XElement? element, string name) =>
        double.TryParse((string?)element?.Attribute(Sfa + name), NumberStyles.Float, CultureInfo.InvariantCulture,
            out double value) && !double.IsNaN(value) && !double.IsInfinity(value) ? value : null;
}
