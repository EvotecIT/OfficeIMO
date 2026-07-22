using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    [Fact]
    public void DocumentReaderEpub_ResolvesChapterReferencesAndProjectsBoundedManifestAssets() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-resources-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithResolvedResources(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            string prefix = result.Source.Path + "::";
            Assert.Contains(result.Links, link =>
                link.Text == "base link" &&
                link.Uri == prefix + "EPUB/text/chapter.xhtml?mode=print#local");
            Assert.Contains(result.Links, link =>
                link.Text == "same fragment" &&
                link.Uri == prefix + "EPUB/text/second.xhtml#second");
            Assert.Contains(result.Links, link =>
                link.Text == "same query" &&
                link.Uri == prefix + "EPUB/text/second.xhtml?view=print");
            Assert.Contains(result.Links, link =>
                link.Text == "root link" &&
                link.Uri == prefix + "EPUB/text/chapter.xhtml#local");
            Assert.DoesNotContain(result.Links, link => link.Text == "unsafe link");
            string markdown = Assert.IsType<string>(result.Markdown);
            Assert.Contains("::EPUB/text/chapter.xhtml?mode=print#local", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/text/second.xhtml#second", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/shared/images/cover%20art.png?display=1#front", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/shared/images/cover%23v2.png", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/shared/audio/chapter.mp3", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/shared/video/clip.mp4#clip", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("../../../outside.png", markdown, StringComparison.Ordinal);

            OfficeDocumentAsset cover = Assert.Single(result.Assets, asset => asset.SourceObjectId == "cover");
            Assert.Equal("image", cover.Kind);
            Assert.Equal(prefix + "EPUB/shared/images/cover art.png", cover.Location.Path);
            Assert.NotNull(cover.PayloadBytes);
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, cover));

            OfficeDocumentAsset reservedImage = Assert.Single(result.Assets, asset => asset.SourceObjectId == "reserved-image");
            Assert.Equal(prefix + "EPUB/shared/images/cover#v2.png", reservedImage.Location.Path);
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, reservedImage));
            Assert.Contains(result.Visuals, visual =>
                visual.Language == "img" &&
                visual.SourceName == prefix + "EPUB/shared/images/cover%23v2.png");

            OfficeDocumentAsset rootImage = Assert.Single(result.Assets, asset => asset.SourceObjectId == "root-image");
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, rootImage));
            Assert.Equal(prefix + "EPUB/shared/images/root.png", rootImage.Location.Path);

            OfficeDocumentAsset audio = Assert.Single(result.Assets, asset => asset.SourceObjectId == "audio");
            Assert.Equal("audio", audio.Kind);
            Assert.Equal("audio/mpeg", audio.MediaType);
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, audio));
            ReaderVisual audioVisual = Assert.Single(result.Visuals, visual =>
                visual.Language == "audio" &&
                visual.SourceName == prefix + "EPUB/shared/audio/chapter.mp3");
            Assert.Equal("audio/mpeg", audioVisual.MimeType);

            OfficeDocumentAsset video = Assert.Single(result.Assets, asset => asset.SourceObjectId == "video");
            Assert.Equal("video", video.Kind);
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, video));
            Assert.Contains(result.Visuals, visual =>
                visual.Language == "video" &&
                visual.SourceName == prefix + "EPUB/shared/video/clip.mp4#clip");
            Assert.Contains(result.Visuals, visual =>
                visual.Language == "img" &&
                visual.SourceName == "data-uri");
            Assert.DoesNotContain(result.Visuals, visual => visual.SourceName == "../../../outside.png");

            Assert.Equal("stylesheet", Assert.Single(result.Assets, asset => asset.SourceObjectId == "styles").Kind);
            Assert.Equal("font", Assert.Single(result.Assets, asset => asset.SourceObjectId == "font").Kind);
            Assert.DoesNotContain(result.Assets, asset => asset.SourceObjectId == "chapter" || asset.SourceObjectId == "second");
            OfficeDocumentAsset remoteImage = Assert.Single(result.Assets, asset => asset.SourceObjectId == "remote-image");
            Assert.Equal("https://cdn.example/remote.png", remoteImage.Location.Path);
            Assert.Null(remoteImage.PayloadBytes);
            Assert.Contains(result.Pages[0].Assets, asset => ReferenceEquals(asset, remoteImage));
            Assert.DoesNotContain(result.Assets, asset => asset.SourceObjectId == "https://cdn.example/remote.png#v2");

            OfficeDocumentDiagnostic[] nonConforming = result.Diagnostics
                .Where(item => item.Code == "epub.reference.non-conforming")
                .ToArray();
            Assert.Equal(2, nonConforming.Length);
            Assert.Contains(nonConforming, item => item.Attributes["reference"] == "/EPUB/shared/images/root.png");
            Assert.Contains(nonConforming, item => item.Attributes["reference"] == "/EPUB/text/chapter.xhtml#local");
            OfficeDocumentDiagnostic unsafeReference = Assert.Single(
                result.Diagnostics,
                item => item.Code == "epub.reference.unsafe" && item.Attributes["reference"] == "../../../outside.png");
            Assert.Equal(OfficeDocumentDiagnosticCategory.Security, unsafeReference.Category);
            Assert.DoesNotContain(result.Assets, asset => asset.SourceObjectId == "../../../outside.png");
            Assert.Contains(result.Diagnostics, item =>
                item.Code == "epub.reference.unsafe" &&
                item.Attributes["reference"] == "../../../outside.xhtml");

            OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(
                OfficeDocumentReadResultJson.Serialize(result, indented: false));
            Assert.Contains(roundTrip.Assets, asset => asset.SourceObjectId == "audio" && asset.Kind == "audio");
            Assert.Contains(roundTrip.Links, link => link.Text == "same fragment" && link.Uri == prefix + "EPUB/text/second.xhtml#second");
            Assert.Contains(roundTrip.Diagnostics, item => item.Code == "epub.reference.unsafe");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_PreservesWindowsVirtualReferencesInMarkdown() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-resources-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithResolvedResources(epubPath);
            using FileStream stream = File.OpenRead(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(
                stream,
                @"C:\books\novel.epub");

            string markdown = Assert.IsType<string>(result.Markdown);
            Assert.Contains("[base link](", markdown, StringComparison.Ordinal);
            Assert.Contains(@"C:\\books\\novel.epub::", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/text/chapter.xhtml?mode=print#local", markdown, StringComparison.Ordinal);
            Assert.Contains("![Cover](", markdown, StringComparison.Ordinal);
            Assert.Contains("::EPUB/shared/images/cover%20art.png?display=1#front", markdown, StringComparison.Ordinal);
            Assert.Contains(result.Links, link =>
                link.Text == "base link" &&
                link.Uri == @"C:\books\novel.epub::EPUB/text/chapter.xhtml?mode=print#local");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }
}
