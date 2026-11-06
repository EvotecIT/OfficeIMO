using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static readonly XNamespace Key = IWorkLegacyXmlSource.Keynote;

    internal static IWorkKeynoteProjection ReadKeynote(IWorkSourceDocument source, IWorkLegacyXmlSource legacy) {
        XElement? slideList = legacy.Root.Element(Key + "slide-list");
        var budget = new IWorkProjectionBudget(source.Options);
        var slides = new List<IWorkKeynoteSlide>();
        if (slideList != null) {
            foreach (XElement slide in slideList.Elements(Key + "slide")) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (slides.Count >= source.Options.MaxProjectedSlides)
                    throw new InvalidDataException("The Keynote '09 presentation exceeds the configured slide limit.");
                slides.Add(ReadKeynoteSlide(source, slide, slides.Count + 1, budget));
            }
        }

        XElement? size = legacy.Root.Element(Key + "size");
        double? width = KeynoteNumber(size, "w");
        double? height = KeynoteNumber(size, "h");
        IWorkCanvasSize? slideSize = width is > 0 && height is > 0 ? new(width.Value, height.Value) : null;
        var diagnostics = new[] {
            LegacyLoss(legacy, "IWORK_LEGACY_KEYNOTE_CONTENT_LIMITED",
                "The Keynote '09 profile recovers ordered slides, direct title and body placeholders, "
                + "and presenter notes. Masters, styles, other drawables, and playback settings are outside the editable projection.")
        };
        return new IWorkKeynoteProjection(source, slides, slideSize, diagnostics, supportsEditableReconstruction: false);
    }

    private static IWorkKeynoteSlide ReadKeynoteSlide(IWorkSourceDocument source, XElement slide,
        int index, IWorkProjectionBudget budget) {
        string name = (string?)slide.Attribute(Key + "name") ?? string.Empty;
        budget.AddTextCharacters(name.Length);
        IWorkTextBox? title = null;
        var boxes = new List<IWorkTextBox>();
        var drawables = new List<IWorkKeynoteDrawable>();
        XElement? page = slide.Element(Key + "page");
        XElement? titlePlaceholder = slide.Element(Key + "title-placeholder");
        XElement? bodyPlaceholder = slide.Element(Key + "body-placeholder");
        XName idAttribute = Sfa + "ID";
        XName referenceAttribute = Sfa + "IDREF";
        string? titleId = (string?)titlePlaceholder?.Attribute(idAttribute);
        string? bodyId = (string?)bodyPlaceholder?.Attribute(idAttribute);
        int availablePlaceholders = (string.IsNullOrEmpty(titleId) ? 0 : 1)
            + (string.IsNullOrEmpty(bodyId) ? 0 : 1);
        if (page != null && availablePlaceholders > 0) {
            var acceptedPlaceholders = new HashSet<XElement>();
            XName titleReference = Sf + "title-placeholder-ref";
            XName bodyReference = Sf + "body-placeholder-ref";
            foreach (XElement drawable in page.Descendants(Sf + "drawables").Elements()) {
                source.CancellationToken.ThrowIfCancellationRequested();
                bool isTitle = drawable.Name == titleReference;
                if (!isTitle && drawable.Name != bodyReference) continue;
                XElement? placeholder = isTitle ? titlePlaceholder : bodyPlaceholder;
                string? expectedId = isTitle ? titleId : bodyId;
                if (placeholder == null || string.IsNullOrEmpty(expectedId)
                    || (string?)drawable.Attribute(referenceAttribute) != expectedId
                    || !acceptedPlaceholders.Add(placeholder)) continue;
                XElement? body = placeholder.Element(Key + "text")
                    ?.Element(Sf + "text-storage")?.Element(Sf + "text-body");
                IWorkTextContent content = ReadText(body, budget, source.CancellationToken);
                if (content.Paragraphs.Any(paragraph => paragraph.Runs.Count > 0)) {
                    var box = new IWorkTextBox(content, ReadPlaceholderGeometry(placeholder), null, null);
                    if (isTitle) title = box;
                    else boxes.Add(box);
                    drawables.Add(new IWorkKeynoteDrawable(box, isTitle));
                }
                if (acceptedPlaceholders.Count == availablePlaceholders) break;
            }
        }

        XElement? notesBody = slide.Element(Key + "notes")
            ?.Element(Sf + "text-storage")?.Element(Sf + "text-body");
        IWorkTextContent notes = ReadText(notesBody, budget, source.CancellationToken);
        return new IWorkKeynoteSlide(index, name, title, boxes, notes, Array.Empty<IWorkImageAsset>(),
            Array.Empty<IWorkTable>(), drawables, isSkipped: false);
    }

    private static IWorkGeometry? ReadPlaceholderGeometry(XElement placeholder) {
        XElement? geometry = placeholder.Element(Sf + "geometry");
        XElement? size = geometry?.Element(Sf + "size");
        XElement? position = geometry?.Element(Sf + "position");
        double? width = KeynoteNumber(size, "w");
        double? height = KeynoteNumber(size, "h");
        double? x = KeynoteNumber(position, "x");
        double? y = KeynoteNumber(position, "y");
        return width is > 0 && height is > 0 && x.HasValue && y.HasValue
            ? new IWorkGeometry(x.Value, y.Value, width.Value, height.Value, 0)
            : null;
    }

    private static double? KeynoteNumber(XElement? element, string name) =>
        double.TryParse((string?)element?.Attribute(Sfa + name), NumberStyles.Float, CultureInfo.InvariantCulture,
            out double value) && !double.IsNaN(value) && !double.IsInfinity(value) ? value : null;
}
