using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static IWorkTextBox? ReadLegacyTextBox(XElement drawable, IWorkLegacyXmlCatalog catalog) {
        XElement? text = drawable.Element(Sf + "text") ?? drawable.Element(Key + "text");
        XElement? body = text?.Element(Sf + "text-storage")?.Element(Sf + "text-body");
        if (body == null) return null;
        IWorkTextContent content = catalog.ReadText(body);
        if (drawable.Element(Sf + "path") != null)
            catalog.AddLoss("IWORK_LEGACY_SHAPE_PAINT_OMITTED",
                "A selected XML shape retains editable text and its frame; its vector path, fill, and outline are omitted.");
        return new IWorkTextBox(content, ReadLegacyGeometry(drawable, catalog), null, null, catalog.Identity(drawable));
    }

    private static IWorkImageAsset? ReadLegacyImage(IWorkSourceDocument source, XElement drawable,
        IWorkProjectionBudget budget, IWorkLegacyXmlCatalog catalog) {
        budget.AddImage();
        XElement? content = drawable.Element(Sf + "content");
        XElement? media = catalog.Resolve(content?.Elements().FirstOrDefault(), "image-media");
        XElement? filtered = catalog.Resolve(media?.Elements().FirstOrDefault(e =>
            e.Name == Sf + "filtered-image" || e.Name == Sf + "filtered-image-ref"), "filtered-image");
        XElement? unfiltered = catalog.Resolve(filtered?.Elements().FirstOrDefault(e =>
            e.Name == Sf + "unfiltered" || e.Name == Sf + "unfiltered-ref"), "unfiltered");
        XElement? data = catalog.Resolve(unfiltered?.Elements().FirstOrDefault(e =>
            e.Name == Sf + "data" || e.Name == Sf + "data-ref"), "data");
        string? path = (string?)data?.Attribute(Sf + "path");
        IWorkPackageEntry? entry = path == null ? null : catalog.PackageEntry(path);
        if (entry == null) return MissingImage();
        string mediaType = IWorkDrawingReader.MediaType(entry.Path, entry.Bytes);
        if (!IWorkDrawingReader.IsEditableOwnerImageMediaType(mediaType)) return MissingImage();
        int? width, height;
        long decodedBytes;
        bool reused = catalog.TryGetValidatedImage(entry, out var metadata);
        if (reused) {
            (width, height) = (metadata.Width, metadata.Height);
            decodedBytes = metadata.DecodedBytes;
        } else {
            (width, height) = IWorkImageInfo.Read(entry.Bytes, mediaType,
                budget.RemainingDecodedImageBytes, out decodedBytes, source.CancellationToken);
        }
        budget.AddDecodedImageBytes(decodedBytes);
        if (!width.HasValue || !height.HasValue) return MissingImage();
        budget.AddProjectedImageBytes(entry.Bytes.LongLength);
        budget.AddTextCharacters(entry.Path.Length);
        if (!reused) catalog.RememberValidatedImage(entry, width.Value, height.Value, decodedBytes);
        bool hasMask = drawable.Element(Sf + "crop-geometry") != null
            || drawable.Element(Sf + "masking-shape-path-source") != null
            || media?.Element(Sf + "alpha-mask-path") != null;
        if (hasMask) catalog.AddLoss("IWORK_LEGACY_IMAGE_MASK_OMITTED",
            "A selected XML image retains its source payload and frame; its crop or mask is outside the editable projection.");
        return new IWorkImageAsset(Path.GetFileName(entry.Path), entry.Path, mediaType, entry.Bytes,
            width, height, ReadLegacyGeometry(drawable, catalog), hasMask, null, null, catalog.Identity(drawable));

        IWorkImageAsset? MissingImage() {
            catalog.AddLoss("IWORK_LEGACY_IMAGE_OMITTED",
                "A selected XML image has a missing, unsupported, or invalid embedded payload and is omitted.");
            return null;
        }
    }

    private static IEnumerable<XElement> LegacyCanvasDrawables(XElement? page) {
        if (page == null) yield break;
        foreach (XElement drawable in page.Elements(Sf + "drawables").Elements()) yield return drawable;
        foreach (XElement drawable in page.Elements(Sf + "layers").Elements(Sf + "layer")
            .Elements(Sf + "drawables").Elements()) yield return drawable;
    }

    private static void LegacyDrawableLoss(IWorkLegacyXmlCatalog catalog) =>
        catalog.AddLoss("IWORK_LEGACY_DRAWABLE_OMITTED",
            "An XML canvas contains an unsupported drawable or transformed group; that object is omitted.");
}
