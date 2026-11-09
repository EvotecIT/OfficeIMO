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

    internal static void RenderPages(PowerPointDocument presentation, OfficeImageExportFormat format,
        PowerPointImageExportOptions options, OfficeImageExportConsumer consumer, CancellationToken cancellationToken) {
        var slides = new List<(int Number, PowerPointSlide Slide)>();
        OfficeImageExportBatchProcessor.RunWithPreflight(options, token => {
            options.Validate();
            if (options.PageSelector == null && options.SlideNumbers?.Any(number => number < 1 || number > presentation.Slides.Count) == true)
                throw new ArgumentOutOfRangeException(nameof(options.SlideNumbers), "Selected slide numbers must exist in the presentation.");
            var selected = options.SlideNumbers == null || options.SlideNumbers.Count == 0
                ? null : new HashSet<int>(options.SlideNumbers);
            IEnumerable<int> numbers = options.PageSelector?.Resolve(presentation.Slides.Count, options.MaximumOutputCount)
                ?? Enumerable.Range(1, presentation.Slides.Count).ToArray();
            foreach (int number in numbers) {
                token.ThrowIfCancellationRequested();
                PowerPointSlide slide = presentation.Slides[number - 1];
                if ((!options.IncludeHiddenSlides && slide.Hidden) || (options.PageSelector == null && selected != null && !selected.Contains(number))) continue;
                slides.Add((number, slide));
            }
            return slides.Count;
        }, (accept, token) => {
            var imageBudget = new PowerPointSlideImageRenderer.PowerPointEmbeddedImageReadBudget(
                options.MaximumEmbeddedImageBytes, options.MaximumTotalEmbeddedImageBytes);
            var snapshots = new List<(int Number, PowerPointSlide Slide, PowerPointSlideVisualInfo Snapshot)>(slides.Count);
            foreach (var slide in slides) {
                token.ThrowIfCancellationRequested();
                snapshots.Add((slide.Number, slide.Slide, PowerPointSlideImageRenderer.CreateSnapshot(slide.Slide, options, token, imageBudget)));
            }
            // The presentation operation shares source limits and findings across its selected pages.
            var diagnostics = snapshots.SelectMany(item => item.Snapshot.Diagnostics).ToArray();
            var budget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);
            for (int index = 0; index < snapshots.Count; index++) {
                token.ThrowIfCancellationRequested();
                var item = snapshots[index];
                var snapshot = new PowerPointSlideVisualInfo(item.Snapshot.Drawing, diagnostics);
                OfficeImageExportResult result = PowerPointSlideImageRenderer.Render(item.Slide, format, options, token, budget, snapshot);
                accept(result.WithIdentity("Slide " + item.Number, "PowerPoint slide " + item.Number).WithSequence(index, snapshots.Count));
            }
        }, consumer, cancellationToken);
    }
}
