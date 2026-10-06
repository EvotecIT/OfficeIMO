using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfOrganizerRasterScaleTests {
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
