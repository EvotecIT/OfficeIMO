using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfIdentifierTests {
    [Theory]
    [InlineData("CorpKeyRecoveryAgent", "Corp\u200CKey\u200CRecovery\u200CAgent")]
    [InlineData("CorpUserV2", "Corp\u200CUser\u200CV\u200C2")]
    [InlineData("CorpEFSV2", "Corp\u200CEFSV\u200C2")]
    public void WithIdentifierBreaks_WrappedWordAndPdf_ExtractsExactIdentifier(string identifier, string display) {
        Assert.Equal(display, WordPaperText.WithIdentifierBreaks(identifier));
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        new WordPaperStyle().SetCellText(table.Rows[0].Cells[0], display);
        table.SetColumnWidthsPercentage(15, 85);
        using MemoryStream word = new();
        document.Save(word);
        word.Position = 0;
        using WordDocument reopened = WordDocument.Load(word);
        Assert.Equal(identifier, WordPaperText.RemoveOptionalBreaks(reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text));
        byte[] pdf = reopened.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        string extracted = PdfReadDocument.Open(pdf).ExtractText().Trim();
        Assert.Equal(identifier, extracted);
        Assert.DoesNotContain('\u200B', extracted);
        string? proof = Environment.GetEnvironmentVariable("OFFICEIMO_IDENTIFIER_PROOF_DIR");
        if (!string.IsNullOrWhiteSpace(proof)) {
            Directory.CreateDirectory(proof);
            File.WriteAllBytes(Path.Combine(proof, identifier + ".pdf"), pdf);
            File.WriteAllBytes(Path.Combine(proof, identifier + ".docx"), word.ToArray());
        }
    }

    [Fact]
    public void CanvasLogicalText_AfterSharedAnchorExtraction_PreservesBoundedSelection() {
        byte[] pdf = PdfDocument.Create().Canvas(canvas => canvas.ActualText("OriginalIdentifier", 20, 20, 120, 12,
            paint => paint.Text("Painted glyphs", 20, 20, 120, 12))).ToBytes();
        Assert.Equal("OriginalIdentifier", PdfReadDocument.Open(pdf).ExtractText().Trim());
    }

    [Theory]
    [InlineData(220)]
    [InlineData(55)]
    public void OptionalBreaks_TaggedHyperlink_ExtractsOneExactIdentifier(double width) {
        const string identifier = "CorpKeyRecoveryAgent";
        const string uri = "https://example.test/template";
        string display = WordPaperText.WithIdentifierBreaks(identifier).Replace('\u200C', '\u200B');
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(width + 40), MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20),
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        }.UseTextFallbacks()).Paragraph(paragraph => paragraph.Link(display, uri)).ToBytes();
        Assert.Equal(identifier, PdfReadDocument.Open(pdf).ExtractText().Trim());
        Assert.Contains(uri, PdfInspector.Inspect(pdf).LinkUris);
        string? proof = Environment.GetEnvironmentVariable("OFFICEIMO_IDENTIFIER_PROOF_DIR");
        if (!string.IsNullOrWhiteSpace(proof)) {
            Directory.CreateDirectory(proof);
            File.WriteAllBytes(Path.Combine(proof, "tagged-link-" + width + ".pdf"), pdf);
        }
    }

    [Fact]
    public void RemoveOptionalBreaks_TextExport_PreservesRealSpacesAndUnicodeJoiners()
        => Assert.Equal("Corp Name\u200D2 فارسی\u200Cمتن", WordPaperText.RemoveOptionalBreaks("Corp\u200B Name\u200D2 فارسی\u200Cمتن"));
}
