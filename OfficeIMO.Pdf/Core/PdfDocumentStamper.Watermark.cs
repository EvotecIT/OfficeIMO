namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentStamper {
    /// <summary>Reads editable settings for watermarks created by this API, with their current page selections.</summary>
    public IReadOnlyList<PdfWatermarkOptions> ReadWatermarks(PdfLoadOptions? readOptions = null) =>
        PdfStamper.ReadWatermarks(_document.ToBytes(), readOptions ?? _document.ReadOptions);
    /// <summary>
    /// Adds or revises a text or image watermark with a centered rotation origin and common opacity,
    /// placement, and page-selection options. Reusing an identifier replaces that watermark on the
    /// requested pages and removes its previous occurrences on other pages. An unchanged foreground
    /// or background setting preserves its position relative to other page content.
    /// </summary>
    public PdfDocument Watermark(PdfWatermarkOptions options, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(options, nameof(options));
        PdfWatermarkOptions snapshot = options.Clone();
        snapshot.Validate();
        return Content((canvas, page) => {
            double x = snapshot.X ?? (page.Width - snapshot.Width) / 2D;
            double y = snapshot.Y ?? (page.Height - snapshot.Height) / 2D;
            if (snapshot.ImageBytes is { } image) {
                canvas.Image(image, x, y, snapshot.Width, snapshot.Height,
                    rotationAngle: -snapshot.RotationDegrees);
            } else {
                canvas.TextBox(snapshot.Text, x, y, snapshot.Width, snapshot.Height,
                    new PdfCanvasTextBoxStyle {
                        Font = snapshot.Font, FontSize = snapshot.FontSize, TextColor = snapshot.Color,
                        Align = PdfAlign.Center, VerticalAlign = PdfVerticalAlign.Middle,
                        PaddingX = 0D, PaddingY = 0D, Background = null, BorderColor = null, BorderWidth = 0D
                    }, -snapshot.RotationDegrees);
            }
        }, new PdfCanvasStampOptions {
            ContentIdentifier = snapshot.Id,
            WatermarkSettings = snapshot,
            TargetPages = snapshot.TargetPages, Opacity = snapshot.Opacity, BehindContent = snapshot.BehindContent
        }, readOptions);
    }
}
