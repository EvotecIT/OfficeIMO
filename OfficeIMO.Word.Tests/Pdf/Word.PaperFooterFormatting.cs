using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPaperFooterFormattingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PageFieldsUseSourceRunTypographyBeforeOrAfterInsertion(bool styleFirst, bool total) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        document.DifferentFirstPage = false;
        document.DifferentOddAndEvenPages = false;
        WordParagraph footer = document.FooterDefaultOrCreate.AddParagraph("Report footer ");
        if (styleFirst) Style(footer);
        footer.AddPageNumber(total, separator: " / ");
        if (!styleFirst) Style(footer);
        document.AddParagraph("First page");
        document.AddPageBreak();
        document.AddParagraph("Second page");

        using var serialized = new MemoryStream();
        document.Save(serialized);
        serialized.Position = 0;
        using WordDocument reopened = WordDocument.Load(serialized);
        var runs = reopened.FooterDefaultOrCreate._footer!.Descendants<Run>().Where(run => run.Elements<Text>().Any()).ToArray();
        Assert.Equal(total ? 4 : 2, runs.Length);
        foreach (Run run in runs) {
            Assert.Equal("17", run.RunProperties?.FontSize?.Val?.Value);
            Assert.Equal("506070", run.RunProperties?.Color?.Val?.Value);
            Assert.Equal("Arial", run.RunProperties?.RunFonts?.Ascii?.Value);
        }
        byte[] bytes = reopened.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        string? proof = Environment.GetEnvironmentVariable("OFFICEIMO_PAPER_PROOF_DIR");
        if (!string.IsNullOrWhiteSpace(proof)) {
            Directory.CreateDirectory(proof);
            string name = "footer-" + styleFirst + "-" + total;
            File.WriteAllBytes(Path.Combine(proof, name + ".pdf"), bytes);
            File.WriteAllBytes(Path.Combine(proof, name + ".docx"), serialized.ToArray());
        }
        PdfReadDocument pdf = PdfReadDocument.Open(bytes, new PdfLoadOptions { IncludeArtifactText = true });
        Assert.Equal(2, pdf.Pages.Count);
        for (int index = 0; index < 2; index++) {
            PdfTextSpan[] spans = pdf.Pages[index].GetTextSpans().Where(span => span.Y < 75 && !string.IsNullOrWhiteSpace(span.Text)).ToArray();
            Assert.True(spans.Length > 0, string.Join(" | ", pdf.Pages[index].GetTextSpans().Select(span => $"{span.Text} @ {span.X},{span.Y} ({span.FontSize})")));
            Assert.Equal("Report footer " + (index + 1) + (total ? " / 2" : ""), string.Concat(spans.Select(span => span.Text)).TrimEnd());
            foreach (PdfTextSpan span in spans) {
                Assert.Equal(8.5D, span.FontSize);
                Assert.Equal("506070", span.Color?.ToRgbHex());
            }
        }
    }

    [Fact]
    public void PageNumberFormattingDoesNotRestyleAnIndependentRun() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Source ");
        WordParagraph independent = paragraph.AddText("Independent ");
        independent.FontSize = 16;
        independent.ColorHex = "AA0000";
        paragraph.AddPageNumber(true);
        Style(paragraph);
        Assert.Equal(16, independent.FontSize);
        Assert.Equal("AA0000", independent.ColorHex);
        Assert.All(paragraph._paragraph.Elements<SimpleField>().SelectMany(field => field.Elements<Run>()),
            run => Assert.Equal("17", run.RunProperties?.FontSize?.Val?.Value));
        paragraph.FontSizePoints = null;
        paragraph.ColorHex = "";
        Assert.All(paragraph._paragraph.Elements<SimpleField>().SelectMany(field => field.Elements<Run>()), run => {
            Assert.Null(run.RunProperties?.FontSize);
            Assert.Null(run.RunProperties?.Color);
        });
    }

    [Fact]
    public void PaperPageChromeUsesTheGenericFieldTypographyContract() {
        using WordDocument document = WordDocument.Create();
        new WordPaperStyle { FontFamily = "Arial", Muted = "506070" }.AddPageChrome(document, "Title", "Product", "Scope");
        document.AddParagraph("Body");
        using var serialized = new MemoryStream();
        document.Save(serialized);
        serialized.Position = 0;
        using WordDocument reopened = WordDocument.Load(serialized);
        foreach (WordFooter footer in new[] { reopened.FooterDefaultOrCreate, reopened.FooterEvenOrCreate }) {
            Assert.All(footer._footer!.Descendants<Run>(), run => {
                Assert.Equal("16", run.RunProperties?.FontSize?.Val?.Value);
                Assert.Equal("506070", run.RunProperties?.Color?.Val?.Value);
            });
        }
    }

    private static void Style(WordParagraph paragraph) {
        paragraph.FontFamily = "Arial";
        paragraph.FontSizePoints = 8.5D;
        paragraph.Color = OfficeColor.ParseHex("506070");
    }
}
