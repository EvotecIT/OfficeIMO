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
        IWorkKeynoteProjection projection = source.ReadKeynote();
        IWorkImportMode mode = source.RequestedImportMode;
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent;
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException("The Keynote source has no supported editable slides.");
        }

        IWorkPreviewAsset? preview = editable ? null : source.PreferredRasterPreview;
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
                slide.AddPictureInches(image, format, 0, 0, 13.333, 7.5);
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
}