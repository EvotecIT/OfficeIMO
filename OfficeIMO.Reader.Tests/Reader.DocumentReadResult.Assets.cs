using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System.Reflection;
using System.Text.Json;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDocumentReadResultAssetTests {
    [Fact]
    public void DocumentReader_ReadAssets_ReturnsEmbeddedWordImageAssets() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        try {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Logo:").AddImage(imagePath);
                document.Save();
            }

            IReadOnlyList<OfficeDocumentAsset> assets = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadAssets(path);

            OfficeDocumentAsset asset = Assert.Single(assets);
            Assert.Equal("word-image-0000", asset.Id);
            Assert.Equal("image", asset.Kind);
            Assert.Equal("image/png", asset.MediaType);
            Assert.Equal("word-image-0000.png", asset.FileName);
            Assert.Equal(path, asset.Location.Path);
            Assert.True(asset.Width > 0);
            Assert.True(asset.Height > 0);
            Assert.True(asset.PayloadHashMatches(out _));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ExtractAssets_FiltersExistingReadResultAssets() {
        var image = new OfficeDocumentAsset {
            Id = "image-asset",
            Kind = "image",
            MediaType = "image/png"
        };
        var preview = new OfficeDocumentAsset {
            Id = "preview-asset",
            Kind = "preview",
            MediaType = "image/svg+xml"
        };
        var result = new OfficeDocumentReadResult {
            Assets = new[] { image, preview }
        };

        IReadOnlyList<OfficeDocumentAsset> assets = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ExtractAssets(result, asset => asset.Kind == "image");

        OfficeDocumentAsset asset = Assert.Single(assets);
        Assert.Same(image, asset);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsEmbeddedWordImageAssets() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        try {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
                WordImage image = document.AddParagraph("Logo:").InsertImage(imagePath);
                image.Description = "Company logo alt text";
                image.Title = "Company logo title";
                document.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            OfficeDocumentAsset asset = Assert.Single(result.Assets);
            Assert.Equal("image", asset.Kind);
            Assert.Equal("image/png", asset.MediaType);
            Assert.Equal(".png", asset.Extension);
            Assert.Equal("word-image-0000.png", asset.FileName);
            Assert.Equal("Company logo alt text", asset.AltText);
            Assert.Equal("Company logo title", asset.Title);
            Assert.Equal(path, asset.Location.Path);
            Assert.Equal("image", asset.Location.SourceBlockKind);
            Assert.Equal("word-image-0000", asset.Location.BlockAnchor);
            Assert.True(asset.Width > 0);
            Assert.True(asset.Height > 0);
            Assert.NotNull(asset.PayloadBytes);
            Assert.True(asset.LengthBytes > 0);
            Assert.Equal(asset.PayloadBytes!.Length, asset.LengthBytes);
            Assert.True(asset.PayloadHashMatches(out string? actualHash));
            Assert.Equal(actualHash, asset.PayloadHash);
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "AssetCount" &&
                entry.Value == "1" &&
                entry.ValueType == "count");

            using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
            JsonElement jsonAsset = jsonDocument.RootElement.GetProperty("assets")[0];
            Assert.Equal("word-image-0000", jsonAsset.GetProperty("id").GetString());
            Assert.Equal("image/png", jsonAsset.GetProperty("mediaType").GetString());
            Assert.Equal("word-image-0000.png", jsonAsset.GetProperty("fileName").GetString());
            Assert.True(jsonAsset.GetProperty("width").GetInt32() > 0);
            Assert.True(jsonAsset.GetProperty("height").GetInt32() > 0);
            Assert.False(jsonAsset.TryGetProperty("payloadBytes", out _));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsPowerPointSlideImageAssetsFromStream() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Intro slide");
            PowerPointPicture picture = slide.AddPicture(imagePath);
            picture.AltText = "Slide logo alt text";
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "deck.pptx");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("powerpoint-slide-0001-image-0000", asset.Id);
        Assert.Equal("image", asset.Kind);
        Assert.Equal("image/png", asset.MediaType);
        Assert.Equal(".png", asset.Extension);
        Assert.Equal("Slide logo alt text", asset.AltText);
        Assert.Equal("deck.pptx", asset.Location.Path);
        Assert.Equal(1, asset.Location.Slide);
        Assert.Equal("image", asset.Location.SourceBlockKind);
        Assert.True(asset.Width > 0);
        Assert.True(asset.Height > 0);
        Assert.True(asset.PayloadHashMatches(out _));

        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Equal(1, page.Number);
        OfficeDocumentAsset pageAsset = Assert.Single(page.Assets);
        Assert.Same(asset, pageAsset);

        Assert.Contains(result.Metadata, entry =>
            entry.Category == "reader.summary" &&
            entry.Name == "AssetCount" &&
            entry.Value == "1" &&
            entry.ValueType == "count");
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsPowerPointAssetsForEachVisibleSlidePlacement() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide first = presentation.AddSlide();
            first.AddPicture(imagePath);
            PowerPointSlide second = presentation.AddSlide();
            second.AddPicture(imagePath);
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "placements.pptx");

        Assert.Equal(2, result.Assets.Count);
        Assert.Equal(new[] { 1, 2 }, result.Assets.Select(asset => asset.Location.Slide ?? -1).ToArray());
        Assert.Equal(2, result.Pages.Count);
        Assert.Single(result.Pages[0].Assets);
        Assert.Single(result.Pages[1].Assets);
        Assert.NotEqual(result.Assets[0].Id, result.Assets[1].Id);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsPowerPointAssetsForEachDuplicateSlidePlacement() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddPicture(imagePath);
            slide.AddPicture(imagePath);
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "duplicate-placement.pptx");

        Assert.Equal(2, result.Assets.Count);
        Assert.All(result.Assets, asset => Assert.Equal(1, asset.Location.Slide));
        Assert.Equal(2, result.Assets.Select(asset => asset.Id).Distinct(StringComparer.Ordinal).Count());

        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Equal(2, page.Assets.Count);
    }

    [Fact]
    public void DocumentReader_ReadDocument_FlagsImageOnlyPowerPointSlideForOcr() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddPicture(imagePath);
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "image-only.pptx");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);
        Assert.Equal(asset.Id, candidate.AssetId);
        Assert.Equal("image", candidate.Kind);
        Assert.Equal(1, candidate.ImageCount);
        Assert.NotNull(candidate.TextBlockCount);
        Assert.Equal(1, candidate.Location.Slide);
        Assert.Contains("Image asset", candidate.Reason, StringComparison.Ordinal);

        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Equal(1, page.Number);
        Assert.Same(asset, Assert.Single(page.Assets));
        Assert.Same(candidate, Assert.Single(page.OcrCandidates));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "ocr-needed" &&
            diagnostic.Location?.Slide == 1);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsExcelWorksheetImageAssets() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.Cell(1, 1, "Logo sheet");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "Logo", altText: "Company logo");
                ExcelSheet otherSheet = document.AddWorksheet("Other");
                otherSheet.Cell(1, 1, "Other sheet");
                otherSheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "OtherLogo", altText: "Other logo");
                document.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel("Images").ReadDocument(path);

            OfficeDocumentAsset asset = Assert.Single(result.Assets);
            Assert.Equal("excel-sheet-0001-image-0000", asset.Id);
            Assert.Equal("image", asset.Kind);
            Assert.Equal("image/png", asset.MediaType);
            Assert.Equal(".png", asset.Extension);
            Assert.Equal("excel-sheet-0001-image-0000.png", asset.FileName);
            Assert.Equal("Company logo", asset.AltText);
            Assert.Equal(path, asset.Location.Path);
            Assert.Equal("Images", asset.Location.Sheet);
            Assert.Equal("image", asset.Location.SourceBlockKind);
            Assert.Equal(1, asset.Width);
            Assert.Equal(1, asset.Height);
            Assert.True(asset.PayloadHashMatches(out _));

            OfficeDocumentPage page = Assert.Single(result.Pages);
            Assert.Equal("Images", page.Name);
            OfficeDocumentAsset pageAsset = Assert.Single(page.Assets);
            Assert.Same(asset, pageAsset);
            Assert.DoesNotContain(result.Assets, candidate => candidate.Location.Sheet == "Other");

            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "AssetCount" &&
                entry.Value == "1" &&
                entry.ValueType == "count");

            using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
            JsonElement jsonAsset = jsonDocument.RootElement.GetProperty("assets")[0];
            Assert.Equal("excel-sheet-0001-image-0000", jsonAsset.GetProperty("id").GetString());
            Assert.Equal("Company logo", jsonAsset.GetProperty("altText").GetString());
            Assert.Equal(1, jsonAsset.GetProperty("width").GetInt32());
            Assert.Equal(1, jsonAsset.GetProperty("height").GetInt32());
            Assert.Equal("Images", jsonAsset.GetProperty("location").GetProperty("sheet").GetString());
            Assert.False(jsonAsset.TryGetProperty("payloadBytes", out _));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsPlainExcelImageAssetsWhenOpenPasswordOptionIsSet() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.Cell(1, 1, "Plain workbook");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "Logo", altText: "Plain logo");
                document.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel("Images").ReadDocument(
                path,
                new ReaderOptions { OpenPassword = "not-used-for-plaintext" });

            OfficeDocumentAsset asset = Assert.Single(result.Assets);
            Assert.Equal("Plain logo", asset.AltText);
            Assert.Equal("Images", asset.Location.Sheet);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_SharesPayloadForDuplicateExcelRelationshipPlacements() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "First", altText: "First");
                sheet.AddImage(3, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "Second", altText: "Second");
                document.Save();
            }
            PointSecondExcelPictureAtFirstImageRelationship(path);

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            OfficeDocumentAsset[] duplicateRelationshipAssets = result.Assets
                .GroupBy(asset => asset.SourceObjectId, StringComparer.Ordinal)
                .Single(group => group.Count() == 2)
                .ToArray();
            Assert.Same(duplicateRelationshipAssets[0].PayloadBytes, duplicateRelationshipAssets[1].PayloadBytes);
            Assert.All(duplicateRelationshipAssets, asset => Assert.True(asset.PayloadHashMatches(out _)));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_RejectsDuplicateExcelRelationshipPlacementsAboveLimit() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "First", altText: "First");
                sheet.AddImage(3, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "Second", altText: "Second");
                document.Save();
            }
            PointSecondExcelPictureAtFirstImageRelationship(path);

            IOException exception = Assert.Throws<IOException>(() => OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(
                path,
                new ReaderOptions { MaxOpenXmlImagePlacementsPerRelationship = 1 }));

            Assert.Contains("MaxOpenXmlImagePlacementsPerRelationship", exception.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_NormalizeOptions_AppliesOpenXmlSafetyDefaultsWhenOptionsAreNull() {
        MethodInfo method = typeof(DocumentReaderEngine).GetMethod("NormalizeOptions", BindingFlags.NonPublic | BindingFlags.Static)!;
        var normalized = (ReaderOptions)method.Invoke(null, new object?[] { null })!;
        var defaults = new ReaderOptions();

        Assert.Equal(defaults.OpenXmlMaxCharactersInPart, normalized.OpenXmlMaxCharactersInPart);
        Assert.Equal(defaults.MaxOpenXmlImageAssets, normalized.MaxOpenXmlImageAssets);
        Assert.Equal(defaults.MaxOpenXmlImagePlacementsPerRelationship, normalized.MaxOpenXmlImagePlacementsPerRelationship);
        Assert.Equal(defaults.MaxOpenXmlImageAssetBytes, normalized.MaxOpenXmlImageAssetBytes);
        Assert.Equal(defaults.MaxOpenXmlImageTotalAssetBytes, normalized.MaxOpenXmlImageTotalAssetBytes);
    }

    [Fact]
    public void DocumentReader_ReadDocument_FiltersExcelAssetsByCaseInsensitiveSheetAndRange() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.Cell(1, 1, "Inside");
                sheet.Cell(5, 5, "Outside");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "InsideLogo", altText: "Inside logo");
                sheet.AddImage(5, 5, png, "image/png", widthPixels: 12, heightPixels: 10, name: "OutsideLogo", altText: "Outside logo");
                document.Save();
            }
            PointSecondExcelPictureAtFirstImageRelationship(path);

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel("data", "A1:B2").ReadDocument(path);

            OfficeDocumentAsset asset = Assert.Single(result.Assets);
            Assert.Equal("Inside logo", asset.AltText);
            Assert.Equal("Data", asset.Location.Sheet);
            Assert.DoesNotContain(result.Assets, candidate => candidate.AltText == "Outside logo");

            OfficeDocumentPage page = Assert.Single(result.Pages, candidate => candidate.Assets.Count > 0);
            Assert.Equal("Data", page.Name);
            Assert.Same(asset, Assert.Single(page.Assets));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void PointSecondExcelPictureAtFirstImageRelationship(string path) {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
        DrawingsPart drawingsPart = document.WorkbookPart!
            .WorksheetParts
            .Select(part => part.DrawingsPart)
            .First(part => part?.WorksheetDrawing != null)!;
        Xdr.Picture[] pictures = drawingsPart.WorksheetDrawing!.Descendants<Xdr.Picture>().ToArray();
        Assert.True(pictures.Length >= 2);

        string firstRelationshipId = pictures[0].BlipFill!.Blip!.Embed!.Value!;
        pictures[1].BlipFill!.Blip!.Embed!.Value = firstRelationshipId;
        drawingsPart.WorksheetDrawing.Save();
    }
}
