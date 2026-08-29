using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Loads a Pages source into the normal editable Word model, using a visual preview only when requested or necessary.</summary>
    public static WordDocument LoadPages(string path, IWorkReadOptions? options = null) =>
        LoadPagesWithReport(path, options).Document;

    /// <summary>Loads a Pages stream into the normal editable Word model, using a visual preview only when requested or necessary.</summary>
    public static WordDocument LoadPages(Stream stream, IWorkReadOptions? options = null) =>
        LoadPagesWithReport(stream, options).Document;

    /// <summary>Loads a Pages source and returns its Word projection, bounded source model, and loss report.</summary>
    public static IWorkPagesLoadResult LoadPagesWithReport(string path, IWorkReadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectPages(IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages, options));
    }

    /// <summary>Loads a Pages stream and returns its Word projection, bounded source model, and loss report.</summary>
    public static IWorkPagesLoadResult LoadPagesWithReport(Stream stream, IWorkReadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectPages(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages, options));
    }

    private static IWorkPagesLoadResult ProjectPages(IWorkSourceDocument source) {
        IWorkPagesProjection projection = source.ReadPages();
        IWorkImportMode mode = source.RequestedImportMode;
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent;
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException("The Pages source has no supported editable content.");
        }

        IWorkPreviewAsset? preview = editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Pages source has no supported editable content or embedded raster preview.");
        }

        WordDocument document = Create();
        try {
            if (editable) {
                foreach (string paragraph in projection.Paragraphs) document.AddParagraph(paragraph);
                foreach (string textBox in projection.TextBoxes) document.AddTextBox(textBox);
                if (projection.Headers.Count > 0 || projection.Footers.Count > 0) {
                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    foreach (string header in projection.Headers) section.Header.Default!.AddParagraph(header);
                    foreach (string footer in projection.Footers) section.Footer.Default!.AddParagraph(footer);
                }
            } else {
                byte[] bytes = preview!.GetBytes();
                using var image = new MemoryStream(bytes, writable: false);
                (double width, double height) = PreviewSize(preview, 600, 780);
                document.AddParagraph().AddImage(image, PreviewFileName(preview), width, height,
                    description: "Visual fallback from the source Pages package");
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            return new IWorkPagesLoadResult(document, source, projection, projection.CreateImportReport(kind, preview));
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static (double Width, double Height) PreviewSize(IWorkPreviewAsset preview,
        double maximumWidth, double maximumHeight) {
        double width = preview.PixelWidth.GetValueOrDefault(800) * 72d / 96d;
        double height = preview.PixelHeight.GetValueOrDefault(1040) * 72d / 96d;
        double scale = Math.Min(1d, Math.Min(maximumWidth / width, maximumHeight / height));
        return (Math.Max(1, width * scale), Math.Max(1, height * scale));
    }

    private static string PreviewFileName(IWorkPreviewAsset preview) =>
        preview.MediaType == "image/png" ? "pages-preview.png" : "pages-preview.jpg";
}