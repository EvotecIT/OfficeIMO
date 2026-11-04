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
    [InlineData(false, "nestedTable")]
    [InlineData(true, "nestedTable")]
    [InlineData(false, "toc")]
    [InlineData(true, "toc")]
    public void SaveAsPdf_UnsupportedNoteTableControlInventoryPreservesAllSourceParts(bool endnote, string profile) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        W.Table table = CreateNoteTable(new[] { new[] { "OUTERCELLTEXT", "OUTERRIGHTTEXT" } }, repeatHeader: false);
        W.TableCell host = table.Elements<W.TableRow>().Single().Elements<W.TableCell>().First();
        W.SdtBlock control = CreateReadOnlyNoteTableControl(profile);
        host.InsertBefore(control, host.Elements<W.Paragraph>().Single());
        note.Append(table);
        source.Settings.UpdateFieldsOnOpen = false;
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement loadedNotes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        W.SdtBlock loadedControl = Assert.Single(loadedNotes.Descendants<W.SdtBlock>());
        Assert.False(document.Settings.UpdateFieldsOnOpen);
        if (profile == "nestedTable") {
            W.Table nested = Assert.Single(loadedControl.Descendants<W.Table>());
            Assert.All(nested.Elements<W.TableRow>(), row => Assert.Null(row.TableRowProperties));
            Assert.All(nested.Descendants<W.TableCell>(), cell => Assert.Null(cell.TableCellProperties));
        } else Assert.False(Assert.Single(loadedControl.Descendants<W.SimpleField>()).Dirty!.Value);
        string[] original = CaptureReadOnlyNoteControlXml(document);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        // The font inventory runs before supported-profile admission. Rejecting
        // this control later must not excuse editing its source while reading it.
        Assert.Equal(original, CaptureReadOnlyNoteControlXml(document));
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string token in new[] { "INITIALNOTE", "OUTERCELLTEXT", "OUTERRIGHTTEXT", profile == "nestedTable" ? "INNERCACHEDTEXT" : "TOCCACHEDTEXT" })
                Assert.Single(Regex.Matches(text, Regex.Escape(token)));
            Assert.Single(result.Warnings, warning => warning.Code == (endnote
                ? "native-endnote-continuation-approximated" : "native-footnote-placement-approximated"));
            if (endnote) Assert.Contains(result.Warnings, warning => warning.Code == "native-endnote-content-approximated");
            Assert.True(result.HasLoss);
            Assert.Equal(original, CaptureReadOnlyNoteControlXml(document));
            Assert.False(document.Settings.UpdateFieldsOnOpen);
        }
        Assert.Empty(document.ValidateDocument());
    }

    private static W.SdtBlock CreateReadOnlyNoteTableControl(string profile) {
        if (profile == "nestedTable") {
            var table = new W.Table(new W.TableProperties(),
                new W.TableGrid(new W.GridColumn { Width = "1800" }),
                new W.TableRow(new W.TableCell(new W.Paragraph(RichNoteRun("INNERCACHEDTEXT")))));
            return new W.SdtBlock(new W.SdtProperties(new W.SdtAlias { Val = "Read-only nested table" }),
                new W.SdtContentBlock(table));
        }
        // This field remains inert. Its cached text is the fallback content;
        // discovery must not mark it dirty or request a future Word refresh.
        var field = new W.SimpleField(RichNoteRun("TOCCACHEDTEXT")) {
            Instruction = " TOC \\o \"1-3\" \\h \\z \\u ", Dirty = false
        };
        return new W.SdtBlock(new W.SdtProperties(new W.SdtContentDocPartObject(
                new W.DocPartGallery { Val = "Table of Contents" }, new W.DocPartUnique())),
            new W.SdtContentBlock(new W.Paragraph(field)));
    }

    private static string[] CaptureReadOnlyNoteControlXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return new[] { main.Document.OuterXml, main.FootnotesPart?.Footnotes?.OuterXml ?? string.Empty,
            main.EndnotesPart?.Endnotes?.OuterXml ?? string.Empty, main.StyleDefinitionsPart!.Styles!.OuterXml,
            main.DocumentSettingsPart!.Settings!.OuterXml };
    }
}
