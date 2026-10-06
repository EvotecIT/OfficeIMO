using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfRasterScaleTests {
    private const long Budget = StudioPdfSecurityPolicy.MaximumRasterPixels;

    [Theory]
    [InlineData(0.95D, 2D, 2D)]
    [InlineData(1D, 2D, 2D)]
    [InlineData(1D, 3D, 3D)]
    [InlineData(1.25D, 2D, 2.5D)]
    [InlineData(1D, 1D, 1D)]
    [InlineData(0.5D, 1D, 0.5D)]
    [InlineData(1D, 1.5D, 1.5D)]
    [InlineData(2D, 2D, 4D)]
    public void LetterPageMatchesDevicePixels(double zoom, double renderScaling, double expected) {
        Assert.Equal(expected, PdfRasterScale.Compose(zoom, renderScaling, 612D, 792D, Budget));
    }

    [Fact]
    public void ZoomBeyondTheBudgetUsesTheLargestFittingScale() {
        double scale = PdfRasterScale.Compose(1D, 3D, 842D, 1191D, Budget);

        Assert.Equal(2.75D, scale);
        Assert.True(Math.Ceiling(842D * scale) * Math.Ceiling(1191D * scale) <= Budget);
        Assert.True(Math.Ceiling(842D * 3D) * Math.Ceiling(1191D * 3D) > Budget);
    }

    [Fact]
    public void OversizedPagesStayWithinTheBudgetBelowTheStep() {
        double scale = PdfRasterScale.Compose(1D, 2D, 14_400D, 14_400D, Budget);

        Assert.Equal(0.2D, scale);
        Assert.True(Math.Ceiling(14_400D * scale) * Math.Ceiling(14_400D * scale) <= Budget);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-2D)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void UnknownRenderScalingIsTreatedAsOneToOne(double renderScaling) {
        Assert.Equal(1D, PdfRasterScale.Compose(1D, renderScaling, 612D, 792D, Budget));
    }

    [Theory]
    [InlineData(1D, 0.18D)]
    [InlineData(2D, 0.5D)]
    [InlineData(3D, 0.75D)]
    public void OrganizerThumbnailFollowsDisplayScaling(double renderScaling, double expected) {
        using var sceneCoordinator = new PageSceneCoordinator((page, _) => Task.FromResult(TestPdfPageScenes.Create(page)));
        using var renderCoordinator = new PageRenderCoordinator((_, _, _) => throw new InvalidOperationException());
        using var page = new PdfOrganizerPageViewModel(1, 612D, 792D, 0, sceneCoordinator, renderCoordinator);

        page.SetRenderScaling(renderScaling);

        Assert.Equal(expected, page.GetRenderScale(TestPdfPageScenes.Create(1, requiresRasterFallback: true)));
    }
}
