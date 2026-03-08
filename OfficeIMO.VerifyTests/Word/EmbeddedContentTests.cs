using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class EmbeddedContentTests : VerifyTestBase {
    private static async Task DoTest(WordDocument document) {
        document.Save();

        var result = await ToVerifyResult(document._wordprocessingDocument);
        await Verifier.Verify(result, GetSettings());
    }

    private static string GetSampleImagePath() {
        return Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "..",
            "OfficeIMO.Tests",
            "Images",
            "Kulek.jpg"));
    }

    private static string CreateTempTextFile(string extension, string content) {
        var filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO.VerifyTests." + Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(filePath, content);
        return filePath;
    }

    [Fact]
    public async Task EmbeddedFragmentAfterDocument() {
        using var document = WordDocument.Create();
        var before = document.AddParagraph("Before");
        document.AddParagraph("After");
        var embedded = document.AddEmbeddedFragmentAfter(before, "<html><body><p>frag</p><p><strong>bold</strong></p></body></html>");

        Assert.Single(document.EmbeddedDocuments);
        Assert.Equal("text/html", embedded.ContentType);
        Assert.Equal("<html><body><p>frag</p><p><strong>bold</strong></p></body></html>", embedded.GetHtml());

        await DoTest(document);
    }

    [Fact]
    public async Task EmbeddedTextAndPictureControlDocument() {
        var textFilePath = CreateTempTextFile(".txt", "alpha" + Environment.NewLine + "beta");

        try {
            using var document = WordDocument.Create();
            var embedded = document.AddEmbeddedDocument(textFilePath, WordAlternativeFormatImportPartType.TextPlain);
            var pictureControl = document.AddParagraph().AddPictureControl(GetSampleImagePath(), 48, 48, "PictureAlias", "PictureTag");
            pictureControl.Tag = "UpdatedPictureTag";

            Assert.Equal("text/plain", embedded.ContentType);
            Assert.Single(document.EmbeddedDocuments);
            Assert.Single(document.PictureControls);
            Assert.Equal("PictureAlias", pictureControl.Alias);
            Assert.Equal("UpdatedPictureTag", pictureControl.Tag);

            await DoTest(document);
        } finally {
            if (File.Exists(textFilePath)) {
                File.Delete(textFilePath);
            }
        }
    }

    [Fact]
    public async Task RemoveEmbeddedDocumentDeletesPart() {
        var htmlFilePath = CreateTempTextFile(".html", "<html><body><p>remove me</p></body></html>");

        try {
            using var document = WordDocument.Create();
            var embedded = document.AddEmbeddedDocument(htmlFilePath, WordAlternativeFormatImportPartType.Html);
            document.AddParagraph("Remaining paragraph");

            Assert.Single(document.EmbeddedDocuments);
            embedded.Remove();

            Assert.Empty(document.EmbeddedDocuments);
            Assert.Empty(document._wordprocessingDocument.MainDocumentPart!.AlternativeFormatImportParts);

            await DoTest(document);
        } finally {
            if (File.Exists(htmlFilePath)) {
                File.Delete(htmlFilePath);
            }
        }
    }
}
