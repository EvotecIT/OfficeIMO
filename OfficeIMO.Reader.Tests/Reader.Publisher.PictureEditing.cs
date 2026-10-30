using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherPictureEditingTests {
    [Fact]
    public void InMemoryPublisherProjectionIncludesCurrentPictureAssetsWithoutReopeningTheSource() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", "Image", "native-image.pub");
        byte[] source = File.ReadAllBytes(path);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherPicture originalPicture = Assert.Single(document.Pages[0].Pictures);
        var payloadOptions = new ReaderPublisherOptions { IncludeImagePayloads = true };
        var hashOptions = new ReaderOptions { ComputeHashes = true };
        OfficeDocumentReadResult original = document.ToOfficeDocumentReadResult("logical.pub", hashOptions, payloadOptions);
        OfficeDocumentAsset originalAsset = Assert.Single(original.Assets);
        byte[] originalBytes = Assert.IsType<byte[]>(originalAsset.PayloadBytes).ToArray();

        byte[] replacement = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        document.ReplacePictureImage(originalPicture.Id, replacement);
        document.SetPictureBounds(originalPicture.Id, 108, 144, 144, 72);
        PublisherPicture currentPicture = Assert.Single(document.Pages[0].Pictures);
        Assert.NotEqual(originalPicture.ImageId, currentPicture.ImageId);
        int currentImageId = Assert.IsType<int>(currentPicture.ImageId);
        string currentAssetId = "publisher-image-" + currentImageId.ToString(CultureInfo.InvariantCulture);

        OfficeDocumentReadResult current = document.ToOfficeDocumentReadResult("logical.pub", hashOptions, payloadOptions);
        OfficeDocumentAsset editedAsset = Assert.Single(current.Assets, asset => asset.Id == currentAssetId);
        Assert.Equal(replacement, editedAsset.PayloadBytes);
        Assert.Equal(replacement.Length, editedAsset.LengthBytes);
        Assert.Equal("image/png", editedAsset.MediaType);
        Assert.Equal(currentImageId.ToString(CultureInfo.InvariantCulture), editedAsset.SourceObjectId);
        Assert.Equal(OfficeDocumentAssetHash.ComputeSha256Hex(replacement), editedAsset.PayloadHash);
        Assert.NotEqual(originalAsset.PayloadHash, editedAsset.PayloadHash);
        Assert.Equal(document.Images.Select(image => "publisher-image-" + image.Id.ToString(CultureInfo.InvariantCulture)),
            current.Assets.Select(asset => asset.Id));
        Assert.Equal(originalBytes, originalAsset.PayloadBytes);
        Assert.Equal(source, File.ReadAllBytes(path));
        Assert.Equal(document.Pages.Count, current.Pages.Count);
        Assert.Null(current.Source.SourceHash);
        Assert.Null(current.Source.LengthBytes);
        Assert.Contains(current.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_LAYOUT_OMITTED");
        foreach (var finding in document.ReadReport.FidelityDiagnostics)
            Assert.Contains(current.Diagnostics, diagnostic => diagnostic.Code == finding.Code && diagnostic.Message == finding.Message);

        OfficeDocumentReadResult withoutPayloads = document.ToOfficeDocumentReadResult("logical.pub", hashOptions);
        OfficeDocumentAsset metadata = Assert.Single(withoutPayloads.Assets, asset => asset.Id == currentAssetId);
        Assert.Equal(replacement.Length, metadata.LengthBytes);
        Assert.Equal("image/png", metadata.MediaType);
        Assert.Null(metadata.PayloadBytes);
        Assert.Null(metadata.PayloadHash);
    }
}
