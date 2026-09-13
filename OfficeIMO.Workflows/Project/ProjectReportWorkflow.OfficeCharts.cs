using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private static void ValidateOfficeOptions(ProjectOfficeReportOptions options) {
        ArgumentNullException.ThrowIfNull(options); ArgumentNullException.ThrowIfNull(options.Images);
        if (!options.IncludeCharts && !options.IncludeDataTables) throw new ArgumentException("Select charts, data tables, or both.", nameof(options));
    }

    private static void AddWordCharts(WordDocument document, ProjectView view, ProjectOfficeReportOptions options, CancellationToken token) {
        int index = 0;
        ExportImages(view, OfficeImageExportFormat.Png, image => {
            token.ThrowIfCancellationRequested();
            var paragraph = document.AddParagraph();
            paragraph.PageBreakBefore = index++ > 0;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            double scale = Math.Min((view.PageWidth - 2 * view.PageMargin) / image.Width,
                (view.PageHeight - 2 * view.PageMargin - 20) / image.Height);
            using var stream = new MemoryStream(image.Bytes);
            paragraph.AddImage(stream, "project-chart-" + index + ".png", image.Width * scale * 96 / 72, image.Height * scale * 96 / 72,
                description: view.Title + " · " + ProjectView.KindTitle(view.Kind) + " · page " + index);
        }, options.Images, token);
        if (options.IncludeDataTables) document.AddPageBreak();
    }

    private static void AddPowerPointCharts(PowerPointPresentation presentation, ProjectView view, ProjectOfficeReportOptions options, CancellationToken token) {
        ExportImages(view, OfficeImageExportFormat.Png, image => {
            token.ThrowIfCancellationRequested();
            if (presentation.Slides.Count >= view.MaxPages) throw new InvalidOperationException("Report exceeds MaxPages.");
            var slide = presentation.AddSlide();
            double scale = Math.Min(view.PageWidth / image.Width, view.PageHeight / image.Height);
            double width = image.Width * scale, height = image.Height * scale;
            using var stream = new MemoryStream(image.Bytes);
            slide.AddPicturePoints(stream, OfficeImageFormat.Png, (view.PageWidth - width) / 2, (view.PageHeight - height) / 2, width, height);
        }, options.Images, token);
    }
}
