using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class ExcelPictureHyperlinkTests {
    [Fact]
    public void IndependentDrawingMlPictureLinkLoadsWithoutChangingNativeTarget() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Pictures", "PictureExternalLink.xlsx");
        using var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        ExcelImage image = document.Sheets.Single().Images.Single();
        Assert.Equal("https://example.org/photos/control?item=1&view=full", image.HyperlinkUri!.OriginalString);
        Assert.Equal(32, image.WidthPixels);
        Assert.Equal(16, image.HeightPixels);
        using var native = SpreadsheetDocument.Open(path, false);
        Assert.Empty(new OpenXmlValidator().Validate(native));
    }

    [Fact]
    public void PictureTargetsSetChangeRemoveAndPreserveMediaAndGeometry() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue));
        ExcelImage first = sheet.AddImage(2, 1, png, "image/png", 32, 16);
        ExcelImage second = sheet.AddImage(4, 1, png, "image/png", 32, 16);
        var shared = new Uri("https://example.org/shared?item=1&view=full");
        first.HyperlinkUri = shared;
        second.HyperlinkUri = shared;
        Assert.Equal(shared, first.HyperlinkUri);
        Assert.Equal(shared, second.HyperlinkUri);
        Assert.Equal(png, first.ToBytes());
        Assert.Equal(32, first.WidthPixels);
        Assert.Equal(16, first.HeightPixels);

        var changed = new Uri("../photos/changed.png", UriKind.Relative);
        first.HyperlinkUri = changed;
        Assert.Equal(changed, first.HyperlinkUri);
        Assert.Equal(shared, second.HyperlinkUri);
        first.HyperlinkUri = null;
        Assert.Null(first.HyperlinkUri);
        Assert.Equal(shared, second.HyperlinkUri);

        using var stream = new MemoryStream();
        document.Save(stream);
        using var reopened = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
        ExcelImage[] images = reopened.Sheets.Single().Images.ToArray();
        Assert.Null(images[0].HyperlinkUri);
        Assert.Equal(shared, images[1].HyperlinkUri);
        Assert.All(images, image => Assert.Equal(png, image.ToBytes()));
        using var native = SpreadsheetDocument.Open(new MemoryStream(stream.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(native));
        DrawingsPart drawing = native.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        Assert.Single(drawing.HyperlinkRelationships);
    }

    [Fact]
    public void ClearingClickLinkPreservesSharedHoverAndUnrelatedRelationships() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        ExcelImage image = sheet.AddImage(1, 1, png, "image/png", 32, 16);
        image.HyperlinkUri = new Uri("https://example.org/shared");
        DrawingsPart drawing = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        Xdr.NonVisualDrawingProperties properties = drawing.WorksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>().Single();
        string id = properties.GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value!;
        properties.AddChild(new A.HyperlinkOnHover { Id = id }, true);
        HyperlinkRelationship unrelated = drawing.AddHyperlinkRelationship(new Uri("https://example.org/unrelated"), true);

        image.HyperlinkUri = null;

        Assert.Contains(drawing.HyperlinkRelationships, relation => relation.Id == id);
        Assert.Contains(drawing.HyperlinkRelationships, relation => relation.Id == unrelated.Id);
        Assert.Equal(png, image.ToBytes());
    }

    [Fact]
    public void StaleAndReadOnlyPictureHandlesCannotMutateLinks() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        ExcelImage original = sheet.AddImage(1, 1, png, "image/png", 32, 16);
        var target = new Uri("https://example.org/current");
        original.HyperlinkUri = target;
        using var stream = new MemoryStream();
        document.Save(stream);
        Assert.Equal(target, original.HyperlinkUri);
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-PictureLink-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            document.Save(path);
            Assert.Throws<InvalidOperationException>(() => original.HyperlinkUri);
            Assert.Throws<InvalidOperationException>(() => original.HyperlinkUri = target);
            Assert.Equal(target, document.Sheets.Single().Images.Single().HyperlinkUri);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
        using var readOnly = ExcelDocument.Load(new MemoryStream(stream.ToArray()), new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        ExcelImage image = readOnly.Sheets.Single().Images.Single();
        Assert.Equal(target, image.HyperlinkUri);
        Assert.Throws<InvalidOperationException>(() => image.HyperlinkUri = null);
        Assert.Equal(target, image.HyperlinkUri);
    }
    [Fact]
    public void RemovedWorksheetPictureCannotBeReboundByName() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        ExcelImage image = sheet.AddImage(1, 1, png, "image/png", 32, 16);
        image.HyperlinkUri = new Uri("https://example.org/old");
        document.RemoveWorksheet("Pictures");
        document.AddWorksheet("Pictures");
        Assert.Throws<InvalidOperationException>(() => image.HyperlinkUri);
        Assert.Throws<InvalidOperationException>(() => image.HyperlinkUri = null);
    }

    [Fact]
    public void FragmentTargetsRemainDistinctAcrossRelationshipReuseAndReopen() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        ExcelImage first = sheet.AddImage(1, 1, png, "image/png", 32, 16);
        ExcelImage second = sheet.AddImage(4, 1, png, "image/png", 32, 16);
        first.HyperlinkUri = new Uri("https://example.org/photos#first");
        second.HyperlinkUri = new Uri("https://example.org/photos#second");
        using var artifact = document.ToStream();
        using var reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));
        ExcelImage[] images = reopened.Sheets[0].Images.ToArray();
        Assert.Equal("https://example.org/photos#first", images[0].HyperlinkUri!.OriginalString);
        Assert.Equal("https://example.org/photos#second", images[1].HyperlinkUri!.OriginalString);
        images[0].HyperlinkUri = null;
        Assert.Equal("https://example.org/photos#second", images[1].HyperlinkUri!.OriginalString);
    }

}
