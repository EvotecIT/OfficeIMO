using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class EpubEmptyImageExportContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_RejectsAnEmptyExtractionBeforeInvokingConsumers(bool asynchronous) {
        byte[] package = OfficeIMO.Shared.Tests.EpubIntegrityFixtures.OneChapter("<p>Too long</p>");
        var document = EpubDocument.Load(new MemoryStream(package), new EpubReadOptions { MaxTotalTextCharacters = 1 });
        Assert.Empty(document.Chapters);
        Assert.Contains(document.Diagnostics, diagnostic => diagnostic.Code == "epub.chapter.text-total-limit");
        var options = new EpubImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true } };
        int accepted = 0;
        if (asynchronous)
            await Assert.ThrowsAsync<InvalidOperationException>(() => document.ExportImagesAsync(OfficeImageExportFormat.Png,
                (_, _) => { accepted++; return Task.CompletedTask; }, options));
        else
            Assert.Throws<InvalidOperationException>(() => document.ExportImages(OfficeImageExportFormat.Png, _ => accepted++, options));
        Assert.Equal(0, accepted);
    }
}
