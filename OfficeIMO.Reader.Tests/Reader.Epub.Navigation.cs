using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    [Fact]
    public void DocumentReaderEpub_ProjectsPackageNavigationMetadataAndRemoteResourcesToJson() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-navigation-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithNavigationMetadata(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            OfficeDocumentMetadataEntry[] rootfiles = result.Metadata
                .Where(item => item.Category == "epub.container.rootfile")
                .ToArray();
            Assert.Equal(2, rootfiles.Length);
            Assert.Equal("false", rootfiles[0].Attributes["isAvailable"]);
            Assert.Equal("false", rootfiles[0].Attributes["isSelected"]);
            Assert.Equal("true", rootfiles[1].Attributes["isAvailable"]);
            Assert.Equal("true", rootfiles[1].Attributes["isSelected"]);

            OfficeDocumentMetadataEntry creator = Assert.Single(
                result.Metadata,
                item => item.Category == "epub.package.metadata" && item.SourceObjectId == "creator");
            Assert.Equal("Reader Author", creator.Value);
            Assert.Equal("aut", creator.Attributes["role"]);
            OfficeDocumentMetadataEntry refinement = Assert.Single(
                result.Metadata,
                item => item.Category == "epub.package.metadata" && item.Name == "file-as");
            Assert.Equal("#creator", refinement.Attributes["refines"]);

            OfficeDocumentMetadataEntry[] toc = result.Metadata
                .Where(item => item.Category == "epub.navigation.toc")
                .ToArray();
            Assert.Equal(2, toc.Length);
            Assert.Equal("1", toc[0].Attributes["depth"]);
            Assert.Equal("2", toc[1].Attributes["depth"]);
            Assert.Equal("Epub3Navigation", toc[1].Attributes["source"]);
            Assert.Equal(result.Source.Path + "::EPUB/chapters/two.xhtml", toc[1].Location!.Path);
            Assert.Equal("details", toc[1].Location.BlockAnchor);
            Assert.Single(result.Metadata, item => item.Category == "epub.navigation.page-list");
            Assert.Equal(2, result.Metadata.Count(item => item.Category == "epub.navigation.landmarks"));

            OfficeDocumentAsset localAsset = Assert.Single(
                result.Assets,
                item => item.SourceObjectId == "local-cover");
            Assert.NotNull(localAsset.PayloadBytes);
            Assert.EndsWith("::EPUB/images/cover.png", localAsset.Location.Path, StringComparison.Ordinal);
            OfficeDocumentAsset remoteAsset = Assert.Single(
                result.Assets,
                item => item.SourceObjectId == "remote-cover");
            Assert.Null(remoteAsset.PayloadBytes);
            Assert.Equal(0, remoteAsset.LengthBytes);
            Assert.Equal("https://cdn.example/remote.png", remoteAsset.Location.Path);
            Assert.Contains(result.Pages[1].Assets, item => ReferenceEquals(item, remoteAsset));
            Assert.Contains(result.Diagnostics, item =>
                item.Code == "epub.resource.remote" &&
                item.Location?.Path == "https://cdn.example/remote.png");
            Assert.Contains(result.Diagnostics, item =>
                item.Code == "epub.navigation.remote-target" &&
                item.Location?.Path == "https://publisher.example/book");

            OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(
                OfficeDocumentReadResultJson.Serialize(result, indented: false));
            OfficeDocumentMetadataEntry roundTripChild = Assert.Single(
                roundTrip.Metadata,
                item => item.Category == "epub.navigation.toc" && item.Attributes["depth"] == "2");
            Assert.Equal(toc[1].Location.Path, roundTripChild.Location!.Path);
            Assert.Equal("details", roundTripChild.Location.BlockAnchor);
            Assert.Single(roundTrip.Assets, item => item.Location.Path == "https://cdn.example/remote.png");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_BuilderPreservesMetadataAndNavigationLimits() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-navigation-limits-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithNavigationMetadata(epubPath);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddEpubHandler(new EpubReadOptions {
                    MaxMetadataItems = 2,
                    MaxNavigationItems = 2,
                    MaxNavigationDepth = 1
                })
                .Build();

            OfficeDocumentReadResult result = reader.ReadDocument(epubPath);

            Assert.Equal(
                "2",
                Assert.Single(result.Metadata, item => item.Id == "epub-metadata-count").Value);
            Assert.Equal(
                "1",
                Assert.Single(result.Metadata, item => item.Id == "epub-toc-item-count").Value);
            Assert.Contains(result.Diagnostics, item => item.Code == "epub.metadata.count-limit");
            Assert.Contains(result.Diagnostics, item => item.Code == "epub.navigation.depth-limit");
            Assert.Contains(result.Diagnostics, item => item.Code == "epub.navigation.item-count-limit");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }
}
