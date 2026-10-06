using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class ExcelPictureHyperlinkTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RangeCopiesPreservePictureTargetsAndIndependentRelationshipLifetime(bool transpose, bool twoCellAnchor) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue));
        ExcelImage original = twoCellAnchor
            ? sheet.AddImageToRange("A1:B2", png, "image/png")
            : sheet.AddImage(1, 1, png, "image/png", 32, 16);
        var target = new Uri("../photos/control.png?item=1&view=full", UriKind.Relative);
        original.HyperlinkUri = target;

        if (transpose) sheet.TransposeRange("A1:B2", "D4");
        else sheet.CopyRange("A1:B2", "D4");

        ExcelImage copy = sheet.Images.Single(image => image.RowIndex == 4 && image.ColumnIndex == 4);
        Assert.Equal(target.OriginalString, copy.HyperlinkUri?.OriginalString);
        Assert.Equal(png, copy.ToBytes());
        Assert.Equal(twoCellAnchor, copy.HasTwoCellAnchor);
        original.HyperlinkUri = null;
        Assert.Equal(target.OriginalString, copy.HyperlinkUri?.OriginalString);

        using var stream = new MemoryStream();
        document.Save(stream);
        using var reopened = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
        ExcelImage reopenedCopy = reopened.Sheets.Single().Images.Single(image => image.RowIndex == 4 && image.ColumnIndex == 4);
        Assert.Equal(target.OriginalString, reopenedCopy.HyperlinkUri?.OriginalString);
        Assert.Equal(png, reopenedCopy.ToBytes());
        using var native = SpreadsheetDocument.Open(new MemoryStream(stream.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(native));
        Assert.Single(native.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.HyperlinkRelationships);
    }

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

    [Fact]
    public void TemplateSheetsRetainSharedPictureLinksHoverAndMediaWithoutIdCollisions() {
        using var document = ExcelDocument.Create();
        ExcelSheet template = document.AddWorksheet("Template");
        template.CellValue(1, 1, "Region {{Name}}");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        var target = new Uri("https://example.org/template#photo");
        ExcelImage first = template.AddImage(3, 1, png, "image/png", 32, 16);
        ExcelImage second = template.AddImage(5, 1, png, "image/png", 32, 16);
        first.HyperlinkUri = target;
        second.HyperlinkUri = target;
        DrawingsPart drawing = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        string originalId = drawing.HyperlinkRelationships.Single().Id;
        drawing.AddHyperlinkRelationship(target, true, "rId1");
        foreach (var properties in drawing.WorksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>()) {
            properties.GetFirstChild<A.HyperlinkOnClick>()!.Id = "rId1";
        }
        drawing.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().First()
            .AddChild(new A.HyperlinkOnHover { Id = "rId1" }, true);
        drawing.DeleteReferenceRelationship(drawing.HyperlinkRelationships.Single(item => item.Id == originalId));
        template.CellValue(8, 1, "Metric"); template.CellValue(8, 2, "Value");
        template.CellValue(9, 1, "Sales"); template.CellValue(9, 2, 10);
        template.AddChartFromRange("A8:B9", 10, 3, 320, 180, ExcelChartType.ColumnClustered, true);
        document.ApplyTemplateSheets("Template", new IDictionary<string, object?>[] {
            new Dictionary<string, object?> { ["Name"] = "North" },
            new Dictionary<string, object?> { ["Name"] = "South" }
        }, (values, index) => (string)values["Name"]!);
        using var artifact = document.ToStream();
        using var reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));
        foreach (ExcelSheet sheet in reopened.Sheets) {
            Assert.Equal(2, sheet.Images.Count());
            Assert.All(sheet.Images, image => {
                Assert.Equal(target.OriginalString, image.HyperlinkUri!.OriginalString);
                Assert.Equal(png, image.ToBytes());
                Assert.Equal(32, image.WidthPixels); Assert.Equal(16, image.HeightPixels);
            });
        }
        using var native = SpreadsheetDocument.Open(new MemoryStream(artifact.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(native));
        foreach (WorksheetPart sheet in native.WorkbookPart!.WorksheetParts) {
            DrawingsPart part = sheet.DrawingsPart!;
            Assert.Equal("rId1", Assert.Single(part.HyperlinkRelationships).Id);
            Assert.Single(part.WorksheetDrawing!.Descendants<A.HyperlinkOnHover>());
            Assert.DoesNotContain(part.Parts, item => item.RelationshipId == "rId1");
            Assert.Single(part.ChartParts);
        }
    }

    [Fact]
    public void TemplateSheetsPreserveDistinctMediaWithPermutedRelationshipIds() {
        byte[] blue = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        byte[] red = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        using var original = ExcelDocument.Create();
        ExcelSheet template = original.AddWorksheet("Template");
        template.CellValue(1, 1, "Region {{Name}}");
        var target = new Uri("https://example.org/template#photo");
        template.AddImage(3, 1, blue, "image/png", 32, 16).HyperlinkUri = target;
        template.AddImage(5, 1, red, "image/png", 32, 16).HyperlinkUri = target;
        using var serialized = original.ToStream();
        using var edited = new MemoryStream();
        serialized.CopyTo(edited);
        edited.Position = 0;
        using (var zip = new ZipArchive(edited, ZipArchiveMode.Update, leaveOpen: true)) {
            const string relationshipsPath = "xl/drawings/_rels/drawing1.xml.rels";
            const string drawingPath = "xl/drawings/drawing1.xml";
            XDocument relationships;
            using (var stream = zip.GetEntry(relationshipsPath)!.Open()) relationships = XDocument.Load(stream);
            var imageRelationships = relationships.Root!.Elements()
                .Where(element => element.Attribute("Type")!.Value.EndsWith("/image", StringComparison.Ordinal)).ToArray();
            Assert.Equal(2, imageRelationships.Length);
            var hyperlink = relationships.Root.Elements().Single(element => element.Attribute("Type")!.Value.EndsWith("/hyperlink", StringComparison.Ordinal));
            var map = new Dictionary<string, string> {
                [imageRelationships[0].Attribute("Id")!.Value] = "rId2",
                [imageRelationships[1].Attribute("Id")!.Value] = "rId1",
                [hyperlink.Attribute("Id")!.Value] = "rId3"
            };
            foreach (var relation in relationships.Root.Elements()) relation.SetAttributeValue("Id", map[relation.Attribute("Id")!.Value]);
            relationships.Root.ReplaceNodes(imageRelationships[0], imageRelationships[1], hyperlink);
            XDocument drawing;
            using (var stream = zip.GetEntry(drawingPath)!.Open()) drawing = XDocument.Load(stream);
            foreach (var attribute in drawing.Descendants().Attributes().Where(attribute =>
                attribute.Name.NamespaceName == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")) {
                if (map.TryGetValue(attribute.Value, out string? id)) attribute.Value = id;
            }
            zip.GetEntry(relationshipsPath)!.Delete();
            using (var stream = zip.CreateEntry(relationshipsPath).Open()) relationships.Save(stream);
            zip.GetEntry(drawingPath)!.Delete();
            using (var stream = zip.CreateEntry(drawingPath).Open()) drawing.Save(stream);
        }
        byte[] bytes = edited.ToArray();
        using (var native = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(native));
            Assert.Equal(new[] { "rId2", "rId1" }, native.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.Parts.Select(item => item.RelationshipId));
        }
        using var document = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(blue, document.Sheets[0].Images.First().ToBytes());
        Assert.Equal(red, document.Sheets[0].Images.Last().ToBytes());
        document.ApplyTemplateSheets("Template", new IDictionary<string, object?>[] {
            new Dictionary<string, object?> { ["Name"] = "North" },
            new Dictionary<string, object?> { ["Name"] = "South" }
        }, (values, index) => (string)values["Name"]!);
        foreach (ExcelSheet sheet in document.Sheets) {
            ExcelImage[] images = sheet.Images.ToArray();
            Assert.Equal(blue, images[0].ToBytes());
            Assert.Equal(red, images[1].ToBytes());
            Assert.All(images, image => Assert.Equal(target.OriginalString, image.HyperlinkUri!.OriginalString));
        }
        using var saved = document.ToStream();
        byte[] clonedBytes = saved.ToArray();
        using (var native = SpreadsheetDocument.Open(new MemoryStream(clonedBytes), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(native));
        }
        using var reopened = ExcelDocument.Load(new MemoryStream(clonedBytes));
        foreach (ExcelSheet sheet in reopened.Sheets) {
            Assert.Equal(blue, sheet.Images.First().ToBytes());
            Assert.Equal(red, sheet.Images.Last().ToBytes());
            Assert.All(sheet.Images, image => Assert.Equal(target.OriginalString, image.HyperlinkUri!.OriginalString));
        }
    }

}
