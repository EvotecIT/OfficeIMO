using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("missing-part", false)]
    [InlineData("missing-part", true)]
    [InlineData("deleted-note", false)]
    [InlineData("deleted-note", true)]
    public void SaveAsPdf_UnresolvedBodyFootnoteRetainsBodyAndLaterEndnoteWithoutInventingPlacementLoss(
        string missingBody, bool columnFallback) {
        string path = Path.Combine(_directoryWithFiles, $"UnresolvedFootnote-{missingBody}-{columnFallback}.docx");
        using (WordDocument source = CreateFootnoteDiagnosticDocument()) {
            if (columnFallback) source.Sections[0].ColumnCount = 2;
            source.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal);
            WordParagraph broken = source.AddParagraph("DANGLINGPREFIX");
            broken.AddFootNote("DELETEDFOOTNOTEBODY").Paragraph.AddText("DANGLINGTAIL");
            source.AddParagraph("ENDNOTEPREFIX").AddEndNote("KEPTENDNOTEBODY").Paragraph.AddText("ENDNOTETAIL");
            source.AddParagraph("FOLLOWINGBODY");
            var main = source.OpenXmlDocument.MainDocumentPart!;
            long id = Assert.Single(broken._paragraph.Descendants<W.FootnoteReference>()).Id!.Value;
            if (missingBody == "missing-part") main.DeletePart(main.FootnotesPart!);
            else Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), note => note.Id?.Value == id).Remove();
            AssertDanglingBodyFootnoteIsInvalid(source);
            source.Save(path);
        }

        string hashBefore = OrdinaryZeroNoteSourceHash(path);
        using WordDocument document = WordDocument.Load(path);
        string[] validationBefore = AssertDanglingBodyFootnoteIsInvalid(document);
        string[] xmlBefore = UnresolvedFootnoteSourceXml(document);
        var result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in new[] { "DANGLINGPREFIX", "DANGLINGTAIL", "ENDNOTEPREFIX", "ENDNOTETAIL", "KEPTENDNOTEBODY", "FOLLOWINGBODY" })
            Assert.Single(Regex.Matches(visible, token));
        Assert.Contains("DANGLINGPREFIXDANGLINGTAIL", visible, StringComparison.Ordinal);
        Assert.Contains("ENDNOTEPREFIX1ENDNOTETAIL", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("DELETEDFOOTNOTEBODY", visible, StringComparison.Ordinal);
        // No ordinary footnote remains to place, even when columns select a fallback profile.
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Equal(validationBefore, AssertDanglingBodyFootnoteIsInvalid(document));
        Assert.Equal(xmlBefore, UnresolvedFootnoteSourceXml(document));
        Assert.Equal(hashBefore, OrdinaryZeroNoteSourceHash(path));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_UnresolvedBodyFootnoteDoesNotSuppressLaterOrdinaryZero(bool columnFallback, bool tableCell) {
        string path = Path.Combine(_directoryWithFiles, $"UnresolvedBeforeOrdinaryZero-{columnFallback}-{tableCell}.docx");
        using (WordDocument source = CreateOrdinaryZeroNoteDocument(endnote: false, explicitNormal: false)) {
            var main = source.OpenXmlDocument.MainDocumentPart!;
            if (columnFallback) source.Sections[0].ColumnCount = 2;
            if (tableCell) {
                W.Paragraph owner = main.Document.Body!.Elements<W.Paragraph>().First();
                WordTable table = source.AddTable(1, 1, WordTableStyle.TableNormal);
                W.Paragraph placeholder = table.Rows[0].Cells[0].Paragraphs[0]._paragraph;
                W.Table sourceTable = placeholder.Ancestors<W.Table>().First();
                sourceTable.Remove(); owner.InsertBeforeSelf(sourceTable); owner.Remove();
                placeholder.InsertBeforeSelf(owner); placeholder.Remove();
            }
            // Reuse the independently validated ordinary-zero shape; never duplicate a special ID0.
            Assert.Empty(source.ValidateDocument());
            W.Footnote ordinary = Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>());
            Assert.Equal(0L, ordinary.Id?.Value);
            Assert.True(ordinary.Type == null || ordinary.Type.Value == W.FootnoteEndnoteValues.Normal);
            main.Document.Body!.PrependChild(new W.Paragraph(
                new W.Run(new W.Text("DANGLINGPREFIX")),
                new W.Run(new W.FootnoteReference { Id = 9999L }),
                new W.Run(new W.Text("DANGLINGTAIL"))));
            AssertDanglingBodyFootnoteIsInvalid(source);
            source.Save(path);
        }

        string hashBefore = OrdinaryZeroNoteSourceHash(path);
        using WordDocument document = WordDocument.Load(path);
        string[] validationBefore = AssertDanglingBodyFootnoteIsInvalid(document);
        string[] xmlBefore = UnresolvedFootnoteSourceXml(document);
        Assert.NotNull(Assert.Single(document.FootNotes, note => note.ReferenceId == 0).Paragraphs);
        var result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in new[] { ZeroNotePrefix, ZeroNoteSuffix, ZeroNoteBody, ZeroNoteFollowingBody })
            Assert.Single(Regex.Matches(visible, token));
        Assert.Contains("DANGLINGPREFIXDANGLINGTAIL", visible, StringComparison.Ordinal);
        Assert.Contains(ZeroNotePrefix + "1" + ZeroNoteSuffix, visible, StringComparison.Ordinal);
        AssertZeroNoteInlineReferencePosition(pdf);
        if (columnFallback) {
            var warning = Assert.Single(result.Warnings, item => item.Code == "native-footnote-placement-approximated");
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
            Assert.True(result.HasLoss);
            // Tolerant conversion and explicitly strict acceptance remain separate operations.
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        } else Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Equal(validationBefore, AssertDanglingBodyFootnoteIsInvalid(document));
        Assert.Equal(xmlBefore, UnresolvedFootnoteSourceXml(document));
        Assert.Equal(hashBefore, OrdinaryZeroNoteSourceHash(path));
    }

    [Fact]
    public void SaveAsPdf_UnresolvedBodyFootnoteCannotTreatContinuationSeparatorAsOrdinary() {
        string path = Path.Combine(_directoryWithFiles, "UnresolvedContinuationFootnote.docx");
        using (WordDocument source = CreateFootnoteDiagnosticDocument()) {
            WordParagraph broken = source.AddParagraph("DANGLINGPREFIX");
            broken.AddFootNote("DELETEDFOOTNOTEBODY").Paragraph.AddText("DANGLINGTAIL");
            source.AddParagraph(ZeroNotePrefix).AddFootNote(ZeroNoteBody).Paragraph.AddText(ZeroNoteSuffix);
            source.AddParagraph(ZeroNoteFollowingBody);
            var main = source.OpenXmlDocument.MainDocumentPart!;
            W.FootnoteReference reference = Assert.Single(broken._paragraph.Descendants<W.FootnoteReference>());
            Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), note => note.Id?.Value == reference.Id!.Value).Remove();
            reference.Id = 0L;
            W.Footnote continuation = Assert.Single(main.FootnotesPart.Footnotes.Elements<W.Footnote>(), note => note.Id?.Value == 0);
            Assert.Equal(W.FootnoteEndnoteValues.ContinuationSeparator, continuation.Type?.Value);
            continuation.Append(new W.Paragraph(new W.Run(new W.Text("SPECIALSTORYONLY"))));
            source.Save(path);
        }

        string hashBefore = OrdinaryZeroNoteSourceHash(path);
        using WordDocument document = WordDocument.Load(path);
        string[] validationBefore = FootnoteValidationBaseline(document);
        string[] xmlBefore = UnresolvedFootnoteSourceXml(document);
        var mainAfterLoad = document.OpenXmlDocument.MainDocumentPart!;
        W.Footnote target = Assert.Single(mainAfterLoad.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), note => note.Id?.Value == 0);
        Assert.Equal(W.FootnoteEndnoteValues.ContinuationSeparator, target.Type?.Value);
        Assert.Single(mainAfterLoad.Document.Descendants<W.FootnoteReference>(), reference => reference.Id?.Value == 0);
        Assert.Null(Assert.Single(document.FootNotes, note => note.ReferenceId == 0).Paragraphs);
        // SDK existence validation may accept a matching ID; its non-Normal target cannot be an ordinary note.
        var result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Single(pdf.GetPages()); // No legitimate continuation story is needed for this short control.
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("DANGLINGPREFIXDANGLINGTAIL", visible, StringComparison.Ordinal);
        Assert.Contains(ZeroNotePrefix + "1" + ZeroNoteSuffix, visible, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(visible, ZeroNoteBody));
        Assert.DoesNotContain("SPECIALSTORYONLY", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("DELETEDFOOTNOTEBODY", visible, StringComparison.Ordinal);
        AssertZeroNoteInlineReferencePosition(pdf);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.Equal(validationBefore, FootnoteValidationBaseline(document));
        Assert.Equal(xmlBefore, UnresolvedFootnoteSourceXml(document));
        Assert.Equal(hashBefore, OrdinaryZeroNoteSourceHash(path));
    }

    private static string[] AssertDanglingBodyFootnoteIsInvalid(WordDocument document) {
        var errors = document.ValidateDocument();
        Assert.NotEmpty(errors);
        Assert.Contains(errors, error => error.NodeName == "footnoteReference");
        return FootnoteValidationBaseline(document);
    }

    private static string[] FootnoteValidationBaseline(WordDocument document) => document.ValidateDocument()
        .Select(error => error.Id + "|" + error.ErrorType + "|" + error.PartUri + "|" + error.Path + "|" + error.NodeName)
        .OrderBy(value => value, StringComparer.Ordinal).ToArray();

    private static string[] UnresolvedFootnoteSourceXml(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        return new[] { main.Document.OuterXml, main.FootnotesPart?.Footnotes?.OuterXml ?? "<absent>",
            main.EndnotesPart?.Endnotes?.OuterXml ?? "<absent>", main.DocumentSettingsPart!.Settings!.OuterXml,
            main.StyleDefinitionsPart!.Styles!.OuterXml };
    }
}
