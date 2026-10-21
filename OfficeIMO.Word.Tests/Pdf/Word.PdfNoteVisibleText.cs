using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_EndnotesRetainVisibleRevisionTextWithoutDeletedOrMovedFromContent() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        W.Paragraph body = source.AddParagraph("BODYREFERENCE")._paragraph;
        AppendNoteRevisionVisibilityProbe(body, "BODY", firstRevisionId: 100);
        source.AddParagraph("ENDREFERENCE").AddEndNote("ENDVISIBLE");
        W.Paragraph note = source._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!
            .Elements<W.Endnote>().Single(item => item.Id?.Value > 0).Elements<W.Paragraph>().Single();
        AppendNoteRevisionVisibilityProbe(note, "END", firstRevisionId: 200);
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string bodyBefore = main.Document.OuterXml;
        string notesBefore = main.EndnotesPart!.Endnotes!.OuterXml;
        Assert.Equal("BODYREFERENCEBODYKEPT", WordParagraph.ReadVisibleText(main.Document.Body!.Elements<W.Paragraph>().First()));
        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            using var pdf = PdfPigDocument.Open(result.ToBytes());
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string visible in new[] { "BODYREFERENCE", "BODYKEPT", "ENDREFERENCE", "ENDVISIBLE", "ENDKEPT" })
                Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, visible));
            foreach (string removed in new[] { "BODYDELETEDSENTINEL", "BODYMOVEDFROMSENTINEL", "ENDDELETEDSENTINEL", "ENDMOVEDFROMSENTINEL" })
                Assert.DoesNotContain(removed, text, StringComparison.Ordinal);
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
        }
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.EndnotesPart.Endnotes.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SaveAsPdf_ComplexFieldFootnoteFallbackRetainsCachedDisplayWithoutInstructionsOrRemovedRevisions() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.AddParagraph("FOOTREFERENCE").AddFootNote("FOOTVISIBLE");
        W.Paragraph note = source._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value > 0).Elements<W.Paragraph>().Single();
        AppendNoteRevisionVisibilityProbe(note, "FOOT", firstRevisionId: 300);
        note.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" PRIVATE FOOTINSTRUCTIONSENTINEL ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("FOOTCACHEDDISPLAY")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string bodyBefore = main.Document.OuterXml;
        string notesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            using var pdf = PdfPigDocument.Open(result.ToBytes());
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            foreach (string visible in new[] { "FOOTREFERENCE", "FOOTVISIBLE", "FOOTKEPT", "FOOTCACHEDDISPLAY" })
                Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, visible));
            foreach (string removed in new[] { "FOOTDELETEDSENTINEL", "FOOTMOVEDFROMSENTINEL", "FOOTINSTRUCTIONSENTINEL", "PRIVATE" })
                Assert.DoesNotContain(removed, text, StringComparison.Ordinal);
            Assert.Single(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
            Assert.True(result.HasLoss);
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        }
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    private static void AppendNoteRevisionVisibilityProbe(W.Paragraph paragraph, string prefix, int firstRevisionId) {
        DateTime revisionDate = new(2026, 10, 20, 0, 0, 0, DateTimeKind.Utc);
        paragraph.Append(new W.DeletedRun(new W.Run(new W.DeletedText(prefix + "DELETEDSENTINEL"))) {
            Id = firstRevisionId.ToString(System.Globalization.CultureInfo.InvariantCulture), Author = "OfficeIMO Tests", Date = revisionDate
        }, new W.MoveFromRun(new W.Run(new W.Text(prefix + "MOVEDFROMSENTINEL"))) {
            Id = (firstRevisionId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), Author = "OfficeIMO Tests", Date = revisionDate
        }, new W.Run(new W.Text(prefix + "KEPT")));
    }
}
