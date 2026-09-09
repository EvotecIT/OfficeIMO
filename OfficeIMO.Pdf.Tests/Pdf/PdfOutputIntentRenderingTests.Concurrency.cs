using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfOutputIntentRenderingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task ConcurrentOpenedRendersPreserveColdOutputAndPageColorProfiles(bool pageProfile, bool unsupportedOutputIntent) {
        byte[] profile = IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents();
        string content = unsupportedOutputIntent
            ? "q 20 0 0 20 10 10 cm BI /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 ID abc EI Q"
            : pageProfile ? "/Cs cs 0.2 0.4 0.8 scn 10 10 20 20 re f" : "0.2 0.4 0.8 rg 10 10 20 20 re f";
        byte[] bytes = BuildPdf(profile, content,
            resources: pageProfile ? "/ColorSpace << /Cs [/ICCBased 6 0 R] >>" : "",
            profileEntries: "/N 3", outputIntents: unsupportedOutputIntent ? "1" : pageProfile ? "[]" : null);
        var display = new PdfPageDisplayOptions { MaximumDimension = 80 };
        byte[] expected = PdfDocument.Load(bytes).Render.DisplayPage(1, display).Bytes!;
        PdfDocument document = PdfDocument.Load(bytes);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.Render.DisplayPage(1, display, cancelled.Token));
        PdfPageRenderResult[] results = await Task.WhenAll(Enumerable.Range(0, 8).Select(index => Task.Run(() =>
            index % 2 == 0 ? document.Render.DisplayPage(1, display) : document.Render.Pages("1",
                new PdfPageRenderOptions { ThumbnailMaxDimension = 80, ContinueOnError = false }).Single())));
        Assert.All(results, result => {
            Assert.True(result.Succeeded);
            Assert.Equal(expected, result.Bytes);
        });
        Assert.Equal(expected, document.Render.DisplayPage(1, display).Bytes);
    }
}
