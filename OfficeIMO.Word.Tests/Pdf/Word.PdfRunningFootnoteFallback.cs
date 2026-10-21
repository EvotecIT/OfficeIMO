using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_MalformedRunningFootnoteInputRetainsActiveLabelAndTailNoteWithoutLeakingDormantVariants() {
        foreach (bool activeReference in new[] { true, false }) {
            using WordDocument source = CreateMalformedRunningFootnoteInput(activeReference);
            AssertMalformedRunningFootnoteReferenceErrors(source, activeReference);
            using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
            AssertMalformedRunningFootnoteReferenceErrors(document, activeReference);
            Assert.False(document.DifferentFirstPage);
            Assert.False(document.DifferentOddAndEvenPages);
            var main = document._wordprocessingDocument.MainDocumentPart!;
            Assert.Empty(main.Document.Descendants<W.FootnoteReference>());
            Assert.Equal(activeReference ? 3 : 2, main.HeaderParts.Sum(part => part.Header.Descendants<W.FootnoteReference>().Count()));
            Assert.Equal(activeReference ? 3 : 2, main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Count(note => note.Id?.Value > 0));
            string notesBefore = main.FootnotesPart.Footnotes.OuterXml;
            string headersBefore = string.Concat(main.HeaderParts.OrderBy(part => part.Uri.ToString()).Select(part => part.Header.OuterXml));
            var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());

            // Reusing the same conversion result must not append fallback notes
            // or warnings again when its PDF is serialized more than once.
            for (int serialization = 0; serialization < 2; serialization++) {
                using var pdf = PdfPigDocument.Open(result.ToBytes());
                Assert.Equal(2, pdf.NumberOfPages);
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("FIRSTBODY", pdf.GetPage(1).Text, StringComparison.Ordinal);
                Assert.Contains("LASTBODY", pdf.GetPage(2).Text, StringComparison.Ordinal);
                foreach (string dormant in new[] { "DORMANTFIRSTREFERENCE", "DORMANTEVENREFERENCE", "DORMANTFIRSTNOTE", "DORMANTEVENNOTE" })
                    Assert.DoesNotContain(dormant, text, StringComparison.Ordinal);

                if (activeReference) {
                    foreach (var page in pdf.GetPages()) {
                        var reference = AnchoredRunningLetters(page, "RUNNINGREFERENCE");
                        var prefixEnd = reference.Last();
                        // The malformed reference uses an explicitly lossy
                        // running-story fallback. Its raised label survives,
                        // but exact inline placement is outside that contract.
                        var marker = Assert.Single(page.Letters, letter => letter.Value == "1" &&
                            letter.PointSize < prefixEnd.PointSize &&
                            letter.StartBaseLine.Y > prefixEnd.StartBaseLine.Y &&
                            letter.StartBaseLine.Y - prefixEnd.StartBaseLine.Y < prefixEnd.PointSize);
                        Assert.Equal("1", marker.Value);
                        Assert.True(marker.PointSize < prefixEnd.PointSize);
                        Assert.True(marker.StartBaseLine.Y > prefixEnd.StartBaseLine.Y);
                        AnchoredRunningLetters(page, "RUNNINGTAIL");
                    }
                    Assert.Single(Regex.Matches(text, "ACTIVERUNNINGNOTE"));
                    var lastPage = pdf.GetPage(2);
                    var note = AnchoredRunningLetters(lastPage, "ACTIVERUNNINGNOTE");
                    Assert.True(AnchoredRunningLetters(lastPage, "LASTBODY")[0].StartBaseLine.Y > note[0].StartBaseLine.Y);
                    var noteMarker = Assert.Single(lastPage.Letters, letter => letter.Value == "1" &&
                        letter.StartBaseLine.Y > note[0].StartBaseLine.Y &&
                        letter.StartBaseLine.Y - note[0].StartBaseLine.Y < note[0].PointSize &&
                        letter.EndBaseLine.X <= note[0].StartBaseLine.X &&
                        note[0].StartBaseLine.X - letter.EndBaseLine.X < note[0].PointSize);
                    Assert.True(noteMarker.PointSize < note[0].PointSize);
                    var warning = Assert.Single(result.Warnings, item => item.Code == "NativeFootnotePlacementApproximated");
                    Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
                    Assert.True(result.HasLoss);
                    Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
                } else {
                    Assert.All(pdf.GetPages(), page => Assert.Contains("ACTIVEHEADER", page.Text, StringComparison.Ordinal));
                    Assert.DoesNotContain(result.Warnings, warning =>
                        warning.Code.IndexOf("Footnote", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        warning.Code.StartsWith("NativeNote", StringComparison.Ordinal));
                }
            }
            Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
            Assert.Equal(headersBefore, string.Concat(main.HeaderParts.OrderBy(part => part.Uri.ToString()).Select(part => part.Header.OuterXml)));
        }
    }

    private static void AssertMalformedRunningFootnoteReferenceErrors(WordDocument document, bool activeReference) {
        var authoredHeaders = document._wordprocessingDocument.MainDocumentPart!.HeaderParts
            .Where(part => part.Header.Descendants<W.FootnoteReference>().Any()).ToArray();
        Assert.Equal(activeReference ? 3 : 2, authoredHeaders.Length);
        Assert.All(authoredHeaders, part => Assert.Single(part.Header.Descendants<W.FootnoteReference>()));
        var errors = document.ValidateDocument();
        Assert.Equal(authoredHeaders.Length, errors.Count);
        Assert.All(errors, error => {
            Assert.Equal("Sem_MissingReferenceElement", error.Id);
            Assert.Equal(OfficeOpenXmlValidationErrorType.Semantic, error.ErrorType);
            Assert.Equal("footnoteReference", error.NodeName);
        });
        Assert.Equal(authoredHeaders.Select(part => part.Uri.ToString()).OrderBy(uri => uri, StringComparer.Ordinal),
            errors.Select(error => Assert.IsType<string>(error.PartUri)).OrderBy(uri => uri, StringComparer.Ordinal));
    }

    private static WordDocument CreateMalformedRunningFootnoteInput(bool activeReference) {
        WordDocument document = CreateFootnoteDiagnosticDocument();
        document.AddHeadersAndFooters();
        document.DifferentFirstPage = true;
        document.DifferentOddAndEvenPages = true;
        WordHeader header = RequireSectionHeader(document, 0, W.HeaderFooterValues.Default);
        if (activeReference) AppendMalformedRunningFootnote(document, header, 1, "RUNNINGREFERENCE", "ACTIVERUNNINGNOTE", " RUNNINGTAIL");
        else header.AddParagraph("ACTIVEHEADER");
        AppendMalformedRunningFootnote(document, RequireSectionHeader(document, 0, W.HeaderFooterValues.First),
            activeReference ? 2 : 1, "DORMANTFIRSTREFERENCE", "DORMANTFIRSTNOTE");
        AppendMalformedRunningFootnote(document, RequireSectionHeader(document, 0, W.HeaderFooterValues.Even),
            activeReference ? 3 : 2, "DORMANTEVENREFERENCE", "DORMANTEVENNOTE");
        // Keep the stored variants and note bodies, but select only the default
        // running story using the canonical section/document activation flags.
        document.DifferentFirstPage = false;
        document.DifferentOddAndEvenPages = false;
        document.AddParagraph("FIRSTBODY").AddBreak(WordBreakType.Page);
        document.AddParagraph("LASTBODY");
        return document;
    }

    private static void AppendMalformedRunningFootnote(WordDocument document, WordHeader header, long id,
        string prefix, string noteText, string? suffix = null) {
        // Header references cannot resolve the main document's FootnotesPart.
        // Construct this malformed input explicitly to test tolerant conversion,
        // without depending on the authoring API accepting an invalid context.
        W.Paragraph paragraph = header.AddParagraph(prefix)._paragraph;
        paragraph.Append(new W.Run(new W.RunProperties(new W.RunStyle { Val = "FootnoteReference" }),
            new W.FootnoteReference { Id = id }));
        if (suffix != null) paragraph.Append(new W.Run(new W.Text(suffix) {
            Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
        }));
        document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Append(new W.Footnote(
            new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "FootnoteText" }),
                new W.Run(new W.RunProperties(new W.RunStyle { Val = "FootnoteReference" }), new W.FootnoteReferenceMark()),
                new W.Run(new W.Text(noteText)))) { Id = id });
    }
}
