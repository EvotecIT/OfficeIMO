using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    [Fact]
    public void DocumentReaderEpub_ProjectsStructuredSemanticsEndToEnd() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-semantics-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithStructuredSemantics(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            string markdown = Assert.IsType<string>(result.Markdown).Replace("\r\n", "\n");
            Assert.Contains("#### Accessible section", markdown, StringComparison.Ordinal);
            Assert.Contains("> Quoted wisdom.", markdown, StringComparison.Ordinal);
            Assert.Contains("C. Third choice", markdown, StringComparison.Ordinal);
            Assert.Contains("| Measure | Value |", markdown, StringComparison.Ordinal);
            Assert.Contains("```csharp", markdown, StringComparison.Ordinal);
            Assert.Contains("Evidence[^note-alpha].", markdown, StringComparison.Ordinal);
            Assert.Contains("[^note-alpha]: Source **detail**.", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("return", markdown, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("[Read details]", markdown, StringComparison.Ordinal);

            OfficeDocumentBlock heading = Assert.Single(
                result.Blocks,
                block => block.Kind == "heading" && block.Text == "Accessible section");
            Assert.Equal(4, heading.Level);
            Assert.Equal("Quoted wisdom.", Assert.Single(result.Blocks, block => block.Kind == "quote").Text);
            Assert.Equal("Console.WriteLine(3);", Assert.Single(result.Blocks, block => block.Kind == "code").Text);
            Assert.Equal("Source detail.", Assert.Single(result.Blocks, block => block.Kind == "footnote").Text);
            Assert.Equal("C.", Assert.Single(result.Blocks, block => block.Kind == "list-item").Marker);
            OfficeDocumentLink accessibleLink = Assert.Single(result.Links, link => link.Text == "Read details");
            Assert.Equal(result.Source.Path + "::OEBPS/chapter.xhtml#detail", accessibleLink.Uri);
            OfficeDocumentAsset cover = Assert.Single(result.Assets, asset => asset.SourceObjectId == "cover");
            Assert.Equal("Accessible cover", cover.AltText);
            Assert.Contains(Assert.Single(result.Pages).Assets, asset => ReferenceEquals(asset, cover));
            Assert.Contains(result.Visuals, visual => visual.Kind == "image" && visual.Content == "Accessible cover");
            Assert.Contains("officeimo.html.accessibility", result.CapabilitiesUsed);
            Assert.Contains("officeimo.html.footnotes", result.CapabilitiesUsed);
            Assert.Contains("officeimo.html.tables", result.CapabilitiesUsed);
            Assert.Contains("officeimo.html.lists", result.CapabilitiesUsed);
            Assert.Contains("officeimo.html.quotes", result.CapabilitiesUsed);
            Assert.Contains("officeimo.html.code", result.CapabilitiesUsed);

            OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(
                EpubReaderAdapter.ReadDocumentJson(epubPath));
            Assert.Equal("Accessible cover", Assert.Single(roundTrip.Assets, asset => asset.SourceObjectId == "cover").AltText);
            Assert.Contains("officeimo.html.footnotes", roundTrip.CapabilitiesUsed);
            Assert.Contains("[^note-alpha]: Source **detail**.", Assert.IsType<string>(roundTrip.Markdown), StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }
}
