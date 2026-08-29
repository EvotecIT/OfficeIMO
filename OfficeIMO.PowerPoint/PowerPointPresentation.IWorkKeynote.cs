using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Loads a Keynote source into the normal editable PowerPoint model, using a visual preview only when requested or necessary.</summary>
    public static PowerPointPresentation LoadKeynote(string path, IWorkReadOptions? options = null) =>
        LoadKeynoteWithReport(path, options).Document;

    /// <summary>Loads a Keynote stream into the normal editable PowerPoint model, using a visual preview only when requested or necessary.</summary>
    public static PowerPointPresentation LoadKeynote(Stream stream, IWorkReadOptions? options = null) =>
        LoadKeynoteWithReport(stream, options).Document;

    /// <summary>Loads a Keynote source and returns its PowerPoint projection, bounded source model, and loss report.</summary>
    public static IWorkKeynoteLoadResult LoadKeynoteWithReport(string path, IWorkReadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectKeynote(IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote, options));
    }

    /// <summary>Loads a Keynote stream and returns its PowerPoint projection, bounded source model, and loss report.</summary>
    public static IWorkKeynoteLoadResult LoadKeynoteWithReport(Stream stream, IWorkReadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectKeynote(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote, options));
    }

    private static IWorkKeynoteLoadResult ProjectKeynote(IWorkSourceDocument source) {
        IWorkImportMode mode = source.RequestedImportMode;
        IWorkPreviewAsset? preview = mode == IWorkImportMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkImportMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Keynote source has no embedded raster preview.");
        }

        IWorkKeynoteProjection projection = source.ReadKeynote();
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent;
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException("The Keynote source has no supported editable slides.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Keynote source has no supported editable slides or embedded raster preview.");
        }

        PowerPointPresentation presentation = Create();
        try {
            if (editable) {
                foreach (IWorkKeynoteSlide sourceSlide in projection.Slides) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Hidden = sourceSlide.IsSkipped;
                    if (sourceSlide.Title.Length > 0) {
                        slide.AddTextBoxInches(sourceSlide.Title, 0.65, 0.45, 12.0, 1.0);
                    }
                    if (sourceSlide.Body.Count > 0) {
                        slide.AddTextBoxInches(string.Join(Environment.NewLine, sourceSlide.Body), 0.85, 1.65, 11.6, 4.9);
                    }
                    if (sourceSlide.PresenterNotes.Length > 0) slide.Notes.Text = sourceSlide.PresenterNotes;
                }
            } else {
                PowerPointSlide slide = presentation.AddSlide();
                using var image = new MemoryStream(preview!.GetBytes(), writable: false);
                OfficeImageFormat format = preview.MediaType == "image/png"
                    ? OfficeImageFormat.Png
                    : OfficeImageFormat.Jpeg;
                (double left, double top, double width, double height) = PreviewLayout(preview);
                slide.AddPictureInches(image, format, left, top, width, height);
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            return new IWorkKeynoteLoadResult(presentation, source, projection, projection.CreateImportReport(kind, preview));
        } catch {
            presentation.Dispose();
            throw;
        }
    }

    private static (double Left, double Top, double Width, double Height) PreviewLayout(IWorkPreviewAsset preview) {
        const double slideWidth = 13.333;
        const double slideHeight = 7.5;
        double pixelWidth = preview.PixelWidth.GetValueOrDefault(16);
        double pixelHeight = preview.PixelHeight.GetValueOrDefault(9);
        double scale = Math.Min(slideWidth / pixelWidth, slideHeight / pixelHeight);
        double width = pixelWidth * scale;
        double height = pixelHeight * scale;
        return ((slideWidth - width) / 2d, (slideHeight - height) / 2d, width, height);
    }
}
