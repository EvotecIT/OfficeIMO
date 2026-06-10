using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System.Text.Json;
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

            IReadOnlyList<OfficeDocumentAsset> assets = DocumentReader.ReadAssets(path);

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

        IReadOnlyList<OfficeDocumentAsset> assets = DocumentReader.ExtractAssets(result, asset => asset.Kind == "image");

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
                document.AddParagraph("Logo:").AddImage(imagePath);
                document.Save();
            }

            OfficeDocumentReadResult result = DocumentReader.ReadDocument(path);

            OfficeDocumentAsset asset = Assert.Single(result.Assets);
            Assert.Equal("image", asset.Kind);
            Assert.Equal("image/png", asset.MediaType);
            Assert.Equal(".png", asset.Extension);
            Assert.Equal("word-image-0000.png", asset.FileName);
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
            slide.AddPicture(imagePath);
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = DocumentReader.ReadDocument(stream, "deck.pptx");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("powerpoint-slide-0001-image-0000", asset.Id);
        Assert.Equal("image", asset.Kind);
        Assert.Equal("image/png", asset.MediaType);
        Assert.Equal(".png", asset.Extension);
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
    public void DocumentReader_ReadDocument_FlagsImageOnlyPowerPointSlideForOcr() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png");
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddPicture(imagePath);
            presentation.Save();
        }

        stream.Position = 0;
        OfficeDocumentReadResult result = DocumentReader.ReadDocument(stream, "image-only.pptx");

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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                sheet.Cell(1, 1, "Logo sheet");
                sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10, name: "Logo", altText: "Company logo");
                document.Save();
            }

            OfficeDocumentReadResult result = DocumentReader.ReadDocument(path, new ReaderOptions { ExcelSheetName = "Images" });

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
}
