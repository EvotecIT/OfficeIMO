using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentImageOwnershipTests {
    [Theory]
    [InlineData("page")]
    [InlineData("master")]
    [InlineData("shape")]
    [InlineData("writer")]
    [InlineData("impress")]
    public void EmbeddingSnapshotsCallerBytesForClonedAndDeduplicatedImages(string consumer) {
        byte[] expected = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        byte[] caller = (byte[])expected.Clone(); OdfDocument document;
        var bounds = new OdfRect(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(10), OdfLength.Points(10));
        if (consumer == "writer") {
            var writer = OdtDocument.Create(); writer.AddParagraph().AddImage(caller, "pixel.png", bounds.Width, bounds.Height); document = writer;
        } else if (consumer == "impress") {
            var presentation = OdpPresentation.Create(); presentation.AddSlide().AddImage(caller, "pixel.png", bounds); document = presentation;
        } else {
            var drawing = OdgDocument.Create(); var page = drawing.AddPage();
            if (consumer == "page") page.Background.SetBitmap(caller);
            else if (consumer == "master") page.MasterBackground.SetBitmap(caller);
            else page.Shapes.AddImage(caller, "pixel.png", bounds);
            var clone = drawing.ClonePage(0);
            if (consumer == "page") clone.Background.SetBitmap(caller);
            document = drawing;
        }
        string path = Assert.Single(document.PackageEntries, entry => entry.StartsWith("Pictures/", StringComparison.Ordinal));
        Array.Clear(caller, 0, caller.Length);
        Assert.Equal(expected, document.GetPackageEntryBytes(path));
        byte[] returned = document.GetPackageEntryBytes(path); Array.Clear(returned, 0, returned.Length);
        Assert.Equal(expected, document.GetPackageEntryBytes(path));
        using var stream = new MemoryStream(document.ToBytes());
        var reopened = OdfDocument.Load(stream); Assert.Equal(expected, reopened.GetPackageEntryBytes(path));
    }
}
