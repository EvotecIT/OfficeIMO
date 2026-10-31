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
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_FallbackNoteStoriesKeepOneSeparatorAndSourceOrderWithinEachKind(bool sectionEnd) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.Sections[0].ColumnCount = 2;
        source.AddEndnoteProperties(position: sectionEnd ? WordEndnotePosition.SectionEnd : WordEndnotePosition.DocumentEnd);

        source.AddParagraph("REFALPHA").AddFootNote("FOOTFIRST");
        source.AddParagraph("REFBETA").AddEndNote("ENDFIRST");
        source.AddParagraph("REFGAMMA").AddFootNote("FOOTSECOND");
        source.AddParagraph("REFDELTA").AddEndNote("ENDSECOND");
        source.AddParagraph("BODYTAIL");
        ConfigureMixedNoteFlowSeparators(source);
        Assert.Empty(source.ValidateDocument());

        // Import the saved DOCX through the ordinary conversion boundary; the
        // two-column fallback must remain explicit rather than become a native
        // page-placement fidelity assertion.
        byte[] input = source.ToBytes();
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string documentBefore = main.Document.OuterXml;
        string footnotesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        string endnotesBefore = main.EndnotesPart!.Endnotes!.OuterXml;
        var references = main.Document.Descendants<W.FootnoteReference>()
            .Select(reference => "OfficeIMO.Footnote." + reference.Id!.Value)
            .Concat(main.Document.Descendants<W.EndnoteReference>()
                .Select(reference => "OfficeIMO.Endnote." + reference.Id!.Value)).ToArray();

        var result = document.ConvertToPdfWithEvidence(FootnoteDiagnosticOptions());
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = Regex.Replace(string.Concat(pdf.GetPages().Select(page => page.Text)), @"\s+", string.Empty);
            foreach (string token in new[] { "FOOTSEPARATOR", "ENDSEPARATOR", "FOOTFIRST", "FOOTSECOND", "ENDFIRST", "ENDSECOND" }) {
                Assert.Single(Regex.Matches(text, Regex.Escape(token)).Cast<Match>());
            }
            string[] order = { "BODYTAIL", "FOOTSEPARATOR", "FOOTFIRST", "FOOTSECOND", "ENDSEPARATOR", "ENDFIRST", "ENDSECOND" };
            for (int i = 1; i < order.Length; i++)
                Assert.True(text.IndexOf(order[i - 1], StringComparison.Ordinal) < text.IndexOf(order[i], StringComparison.Ordinal), text);
            // Grouping note stories must not alter the authored reference
            // positions or the independent numbering sequences.
            foreach (string marker in new[] { "REFALPHA1", "REFBETAi", "REFGAMMA2", "REFDELTAii" })
                Assert.Contains(marker, text, StringComparison.Ordinal);
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            foreach (string destinationName in references) {
                PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations, item => item.Name == destinationName);
                PdfLinkAnnotation link = Assert.Single(navigation.LinkAnnotations, item => item.DestinationName == destinationName);
                Assert.True(link.IsNamedDestinationLink);
                Assert.True(link.Width > 0D && link.Height > 0D);
                Assert.NotNull(destination.PageNumber);
                Assert.InRange(destination.PageNumber.Value, 1, pdf.NumberOfPages);
            }
        }
        Assert.Single(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-note-flow-separator-omitted");
        Assert.Equal(documentBefore, main.Document.OuterXml);
        Assert.Equal(footnotesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Equal(endnotesBefore, main.EndnotesPart.Endnotes.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    private static void ConfigureMixedNoteFlowSeparators(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Footnote foot = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
            .Single(note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        SetSeparatorText(foot, "FOOTSEPARATOR");
        W.Endnote? end = main.EndnotesPart!.Endnotes!.Elements<W.Endnote>()
            .SingleOrDefault(note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        if (end == null) {
            end = new W.Endnote { Id = -1, Type = W.FootnoteEndnoteValues.Separator };
            main.EndnotesPart.Endnotes.PrependChild(end);
        }
        SetSeparatorText(end, "ENDSEPARATOR");
        W.Settings settings = main.DocumentSettingsPart!.Settings!;
        W.FootnoteDocumentWideProperties? footProperties = settings.GetFirstChild<W.FootnoteDocumentWideProperties>();
        if (footProperties == null) {
            footProperties = new W.FootnoteDocumentWideProperties();
            settings.AddChild(footProperties, true);
        }
        if (!footProperties.Elements<W.FootnoteSpecialReference>().Any(reference => reference.Id?.Value == foot.Id!.Value))
            footProperties.Append(new W.FootnoteSpecialReference { Id = (int)foot.Id!.Value });
        W.EndnoteDocumentWideProperties? endProperties = settings.GetFirstChild<W.EndnoteDocumentWideProperties>();
        if (endProperties == null) {
            endProperties = new W.EndnoteDocumentWideProperties();
            settings.AddChild(endProperties, true);
        }
        if (!endProperties.Elements<W.EndnoteSpecialReference>().Any(reference => reference.Id?.Value == end.Id!.Value))
            endProperties.Append(new W.EndnoteSpecialReference { Id = (int)end.Id!.Value });

        static void SetSeparatorText(W.FootnoteEndnoteType note, string text) {
            note.RemoveAllChildren();
            note.Append(new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact
            }), new W.Run(new W.Text(text))));
        }
    }
}
