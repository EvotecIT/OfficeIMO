using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

internal static class TestPdfPageScenes {
    private static readonly Lazy<byte[]> FixtureBytes = new(() => File.ReadAllBytes(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf")));

    internal static PdfPageScene Create(
        int pageNumber = 1,
        bool requiresRasterFallback = false,
        OfficeDrawing? drawing = null,
        IReadOnlyList<string>? diagnostics = null) {
        drawing ??= new OfficeDrawing(612D, 792D);
        PdfPageInteractionMap interactions = PdfPageInteractionMap.Create(FixtureBytes.Value, pageNumber: 1);
        return new PdfPageScene(
            pageNumber,
            drawing,
            interactions,
            diagnostics ?? Array.Empty<string>(),
            requiresRasterFallback);
    }
}
