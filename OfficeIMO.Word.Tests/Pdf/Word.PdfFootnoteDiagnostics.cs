using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("control")]
    public void SaveAsPdf_ReferencedFootnoteFlowFallbackReportsApproximationOnce(string context) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        // Multiple columns remain outside the qualified page-footnote profile;
        // exercise the real fallback without treating supported notes as lossy.
        source.Sections[0].ColumnCount = 2;
        WordParagraph paragraph = context == "table"
            ? source.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0].SetText("BODY")
            : source.AddParagraph("BODY");
        paragraph.AddFootNote("FOOTNOTEALPHA");
        paragraph.AddFootNote("FOOTNOTEBETA");
        if (context == "control") {
            paragraph._paragraph.InsertBeforeSelf(new W.SdtBlock(
                new W.SdtContentBlock(paragraph._paragraph.CloneNode(true))));
            paragraph._paragraph.Remove();
        }
        source.AddParagraph("LASTBODY");

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        Assert.Equal(2, document._wordprocessingDocument.MainDocumentPart!.Document
            .Descendants<W.FootnoteReference>().Count());
        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        foreach (string token in new[] { "BODY", "LASTBODY", "FOOTNOTEALPHA", "FOOTNOTEBETA" })
            Assert.Contains(token, text, StringComparison.Ordinal);

        // Both notes and duplicate traversal wrappers share one placement
        // approximation for this conversion; their text is retained.
        var warning = Assert.Single(result.Warnings,
            item => item.Code == "NativeFootnotePlacementApproximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.True(result.HasLoss);
        Assert.True(result.Report.HasLoss);
        Assert.Equal(PdfConversionFidelityStatus.Degraded, result.Report.FidelityStatus);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_StoredUnreferencedFootnotesDoNotReportAPlacementApproximation(bool storedOrdinaryNote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        WordParagraph paragraph = source.AddParagraph("BODY");
        if (storedOrdinaryNote) {
            paragraph.AddFootNote("STOREDONLYNOTE");
            Assert.Single(paragraph._paragraph.Descendants<W.FootnoteReference>()).Parent!.Remove();
        }

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        Assert.Empty(main.Document.Descendants<W.FootnoteReference>());
        W.Footnote[] stored = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().ToArray();
        Assert.Contains(stored, note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        Assert.Contains(stored, note => note.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator);
        Assert.Equal(storedOrdinaryNote ? 1 : 0, stored.Count(note => note.Id?.Value > 0));

        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BODY", text, StringComparison.Ordinal);
        Assert.DoesNotContain("STOREDONLYNOTE", text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Warnings,
            warning => warning.Code == "NativeFootnotePlacementApproximated");
    }

    [Fact]
    public void SaveAsPdf_EndnoteOnlyDoesNotReportAFootnotePlacementApproximation() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.AddParagraph("BODY").AddEndNote("ENDNOTEBODY");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        Assert.Empty(document._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.FootnoteReference>());
        Assert.Single(document._wordprocessingDocument.MainDocumentPart.Document.Descendants<W.EndnoteReference>());

        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BODY", text, StringComparison.Ordinal);
        Assert.Contains("ENDNOTEBODY", text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Warnings,
            warning => warning.Code == "NativeFootnotePlacementApproximated");
    }

    private static WordDocument CreateFootnoteDiagnosticDocument() {
        WordDocument document = WordDocument.Create();
        var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" })));
        return document;
    }

    private static WordToPdfOptions FootnoteDiagnosticOptions() => new() {
        IncludePageNumbers = false,
        ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
