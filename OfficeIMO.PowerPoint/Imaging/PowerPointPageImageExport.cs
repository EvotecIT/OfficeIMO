using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Selects slides and applies one encoding budget over the native Drawing scene.</summary>
internal static class PowerPointPageImageExport {
    internal static OfficeImageExportResult Render(PowerPointSlide slide, OfficeImageExportFormat format,
        PowerPointImageExportOptions options, CancellationToken cancellationToken) =>
        PowerPointSlideImageRenderer.Render(slide, format, options, cancellationToken);

    internal static PowerPointSlideVisualInfo CreateSnapshot(PowerPointSlide slide, PowerPointImageExportOptions options) =>
        PowerPointSlideImageRenderer.CreateSnapshot(slide, options);

    internal static void RenderPages(PowerPointPresentation presentation, OfficeImageExportFormat format,
        PowerPointPresentationImageExportOptions options, OfficeImageExportConsumer consumer, CancellationToken cancellationToken) {
        OfficeImageExportExecutionScope.Run(options, cancellationToken, token => {
            options.Validate();
            if (options.SlideNumbers?.Any(number => number < 1 || number > presentation.Slides.Count) == true)
                throw new ArgumentOutOfRangeException(nameof(options.SlideNumbers), "Selected slide numbers must exist in the presentation.");
            var selected = options.SlideNumbers == null || options.SlideNumbers.Count == 0
                ? null : new HashSet<int>(options.SlideNumbers);
            var imageBudget = new PowerPointSlideImageRenderer.PowerPointEmbeddedImageReadBudget(
                options.MaximumEmbeddedImageBytes, options.MaximumTotalEmbeddedImageBytes);
            var snapshots = new List<(int Number, PowerPointSlide Slide, PowerPointSlideVisualInfo Snapshot)>();
            for (int index = 0; index < presentation.Slides.Count; index++) {
                token.ThrowIfCancellationRequested();
                int number = index + 1;
                PowerPointSlide slide = presentation.Slides[index];
                if ((!options.IncludeHiddenSlides && slide.Hidden) || (selected != null && !selected.Contains(number))) continue;
                snapshots.Add((number, slide, PowerPointSlideImageRenderer.CreateSnapshot(slide, options, token, imageBudget)));
            }
            // The presentation operation shares source limits and findings across its selected pages.
            var diagnostics = snapshots.SelectMany(item => item.Snapshot.Diagnostics).ToArray();
            var budget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);
            for (int index = 0; index < snapshots.Count; index++) {
                token.ThrowIfCancellationRequested();
                var item = snapshots[index];
                var snapshot = new PowerPointSlideVisualInfo(item.Snapshot.Drawing, diagnostics);
                OfficeImageExportResult result = PowerPointSlideImageRenderer.Render(item.Slide, format, options, token, budget, snapshot);
                consumer(result.WithIdentity("Slide " + item.Number, "PowerPoint slide " + item.Number).WithSequence(index, snapshots.Count));
            }
            return 0;
        });
    }
}
