using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "image")]
    [InlineData(true, "image")]
    [InlineData(false, "fixed")]
    [InlineData(true, "fixed")]
    [InlineData(false, "cantSplit")]
    [InlineData(true, "cantSplit")]
    [InlineData(false, "rowCantSplit")]
    [InlineData(true, "rowCantSplit")]
    [InlineData(false, "header")]
    [InlineData(true, "header")]
    [InlineData(false, "footer")]
    [InlineData(true, "footer")]
    [InlineData(false, "caption")]
    [InlineData(true, "caption")]
    public void SaveAsPdf_UnsupportedMappedNoteTableRetainsTextAndReportsFallback(bool endnote, string profile) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        W.Table table = CreateVerticallyMergedNoteTable(repeatHeader: profile == "header");
        note.Append(table);
        W.TableRow[] rows = table.Elements<W.TableRow>().ToArray();
        if (profile == "image") {
            W.Paragraph anchor = rows[0].Elements<W.TableCell>().First().Elements<W.Paragraph>().Single();
            using var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng((byte)0, (byte)180, (byte)80));
            new WordParagraph(source, anchor).AddText(string.Empty).AddImage(pixels, "merged-note.png", 24, 16);
        } else if (profile == "fixed") {
            rows[1].TableRowProperties = new W.TableRowProperties(new W.TableRowHeight { Val = 500U, HeightType = W.HeightRuleValues.Exact });
        } else if (profile == "cantSplit") {
            // All rows use the table's effective false default; there is no
            // mixed per-row override list for admission to rely upon.
            foreach (W.TableRow row in rows) row.TableRowProperties = new W.TableRowProperties(new W.CantSplit());
        } else if (profile == "rowCantSplit") {
            rows[1].TableRowProperties = new W.TableRowProperties(new W.CantSplit());
        } else if (profile == "header") {
            // The span is wholly inside two header rows, so the Word owner
            // retains that role rather than suppressing a crossed boundary.
            rows[1].TableRowProperties = new W.TableRowProperties(new W.TableHeader());
        }
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement notes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        string original = notes.OuterXml;
        WordPdfExportOptions options = FootnoteDiagnosticOptions();
        if (profile is "footer" or "caption") options.PdfOptions = new PdfCreateOptions {
            DefaultTableStyle = new PdfTableStyle {
                FooterRowCount = profile == "footer" ? 2 : 0,
                Caption = profile == "caption" ? "Configured table caption" : null
            }
        };
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(options);
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string token in new[] { "INITIALNOTE", "MERGEDANCHOR", "AFTERMERGE", "TOPRIGHT", "MIDRIGHT", "BOTTOMRIGHT" })
                Assert.Single(Regex.Matches(text, token));
            Assert.Single(result.Warnings, warning => warning.Code == (endnote
                ? "native-endnote-continuation-approximated" : "native-footnote-placement-approximated"));
            if (endnote) Assert.Contains(result.Warnings, warning => warning.Code == "native-endnote-content-approximated");
            if (profile == "image" && !endnote)
                Assert.Contains(result.Warnings, warning => warning.Code == "native-footnote-content-omitted");
            Assert.True(result.HasLoss);
            Assert.Equal(original, notes.OuterXml);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NoteTableFlexibleTextMergeUsesEffectiveBodyRoleAcrossHeaderBoundary(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        W.Table table = CreateVerticallyMergedNoteTable(repeatHeader: true);
        note.Append(table);
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement notes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        string original = notes.OuterXml;
        WordPdfExportOptions options = FootnoteDiagnosticOptions();
        options.PdfOptions = new PdfCreateOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers };
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(options);
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string token in new[] { "MERGEDANCHOR", "AFTERMERGE", "TOPRIGHT", "MIDRIGHT", "BOTTOMRIGHT" })
                Assert.Single(Regex.Matches(text, token));
            Assert.DoesNotContain(result.Warnings, warning => warning.Code is
                "native-footnote-placement-approximated" or "native-endnote-continuation-approximated" or "native-endnote-content-approximated");
            PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
            Assert.Contains(tagged.StructureElements, element => element.StructureType == "Table");
            Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "TH");
            Assert.Equal(original, notes.OuterXml);
        }
    }

    [Fact]
    public void SaveAsPdf_TableInSelectedFootnoteSeparatorUsesReportedFallback() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        AddRichNoteFixture(source, endnote: false);
        SetPageFootnoteAdmissionKeepNext(source, "separator", inherited: false, directOverride: false);
        var main = source._wordprocessingDocument.MainDocumentPart!;
        W.Footnote separator = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
            .Single(item => item.Type?.Value == W.FootnoteEndnoteValues.Separator);
        separator.RemoveAllChildren();
        separator.Append(CreateNoteTable(new[] { new[] { "SEPARATORTABLE", "SEPARATOROTHER" } }, repeatHeader: false));
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string original = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.OuterXml;
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Contains("INITIALNOTE", string.Concat(pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        Assert.Single(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Contains(result.Warnings, warning => warning.Code == "native-note-flow-separator-omitted");
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.FootnotesPart.Footnotes.OuterXml);
    }

    private static W.Table CreateVerticallyMergedNoteTable(bool repeatHeader) {
        W.Table table = CreateNoteTable(new[] {
            new[] { "MERGEDANCHOR", "TOPRIGHT" }, new[] { "", "MIDRIGHT" }, new[] { "AFTERMERGE", "BOTTOMRIGHT" }
        }, repeatHeader);
        W.TableRow[] rows = table.Elements<W.TableRow>().ToArray();
        rows[0].Elements<W.TableCell>().First().TableCellProperties!.VerticalMerge = new W.VerticalMerge { Val = W.MergedCellValues.Restart };
        rows[1].Elements<W.TableCell>().First().TableCellProperties!.VerticalMerge = new W.VerticalMerge { Val = W.MergedCellValues.Continue };
        return table;
    }
}
