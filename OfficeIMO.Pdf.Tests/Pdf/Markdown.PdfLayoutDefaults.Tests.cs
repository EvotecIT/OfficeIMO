using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using System.IO;
using System.Linq;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownPdfLayoutDefaultsTests {
    private static double StartX(Letter[] letters, string word) {
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        Assert.True(index >= 0, "Expected '" + word + "' in '" + text + "'.");
        return letters[index].StartBaseLine.X;
    }

    private static Letter[] Letters(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        return pdf.GetPages().SelectMany(page => page.Letters).Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
    }

    [Fact]
    public void NestedListsAndCodeKeepTheirListItemIndent() {
        const string markdown = "Bodytext paragraph.\n\n1. Firststep\n\n   ```\n   codeline\n   ```\n\n2. Secondstep\n   - Subbullet\n     - Deepbullet\n";

        byte[] bytes = MarkdownDocument.Parse(markdown).ToPdf().ToBytes();

        Letter[] letters = Letters(bytes);
        double body = StartX(letters, "Bodytext");
        double first = StartX(letters, "Firststep");
        double code = StartX(letters, "codeline");
        double sub = StartX(letters, "Subbullet");
        double deep = StartX(letters, "Deepbullet");
        Assert.True(code >= first - 1, "Code starts at " + code + ", item text at " + first + ".");
        Assert.True(sub > first + 5, "Sub-bullet text starts at " + sub + ", item text at " + first + ".");
        Assert.True(deep > sub + 5, "Third level starts at " + deep + ", second level at " + sub + ".");
        Assert.True(first > body + 10);
    }

    [Fact]
    public void InlineImagesWrittenAsTextAreReported() {
        MarkdownDocument document = MarkdownDocument.Parse("Status ![Build](https://example.test/build.svg) badge.");

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new MarkdownPdfExportOptions());

        Assert.Contains(result.Warnings, warning => warning.Code == "InlineImageAsText" && warning.Source == "https://example.test/build.svg");
    }

    [Fact]
    public void LoadedReadme_ShowsItsOwnRelativeImagesWithoutExtraOptions() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-md-readme-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            File.WriteAllBytes(Path.Combine(folder, "logo.png"), OfficeIMO.Drawing.OfficePngWriter.EncodeRgba(2, 2, new byte[16]));
            File.WriteAllText(Path.Combine(folder, "README.md"), "# Readme\n\n![Logo](logo.png)\n");

            PdfCore.PdfDocumentConversionResult result = MarkdownDocument.Load(Path.Combine(folder, "README.md")).ConvertToPdfWithEvidence();
            byte[] bytes = result.ToBytes();

            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "LocalImageDisabled");
            Assert.NotEmpty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void LeadInParagraphStaysWithItsCodeBlock() {
        string filler = string.Join("\n\n", Enumerable.Range(0, 34).Select(index => "Filler paragraph " + index + "."));
        string markdown = filler + "\n\nBuild and preview:\n\n```\nline one\nline two\nline three\nline four\nline five\nline six\n```\n";

        byte[] bytes = MarkdownDocument.Parse(markdown).ToPdf().ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        int leadInPage = pdf.GetPages().First(page => page.Text.Contains("Build and preview:")).Number;
        int codeStartPage = pdf.GetPages().First(page => page.Text.Contains("line one")).Number;
        int codeEndPage = pdf.GetPages().First(page => page.Text.Contains("line six")).Number;
        Assert.Equal(leadInPage, codeStartPage);
        Assert.Equal(codeStartPage, codeEndPage);
    }
}
