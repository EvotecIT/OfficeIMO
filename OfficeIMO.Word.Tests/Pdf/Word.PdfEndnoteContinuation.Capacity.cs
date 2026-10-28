using System.Globalization;
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
    public void SaveAsPdf_NoNoticeEndnoteContinuationCapacityUsesFullBodyFrame() {
        const int pageHeight = 500, margin = 40, bodyAdvance = 20, firstSeparatorAdvance = 20;
        const int shortSeparatorAdvance = 12, tallSeparatorAdvance = 24, paragraphCount = 65;
        int frameHeight = pageHeight - 2 * margin;
        int firstCount = (frameHeight - bodyAdvance - firstSeparatorAdvance) / bodyAdvance;
        int continuedCount = frameHeight / bodyAdvance;
        int[] expectedCounts = { firstCount, continuedCount, continuedCount, paragraphCount - firstCount - 2 * continuedCount };
        using WordDocument shorter = CreateEndnoteCapacityDocument(shortSeparatorAdvance, paragraphCount, out string[] tokens);
        using WordDocument taller = CreateEndnoteCapacityDocument(tallSeparatorAdvance, paragraphCount, out string[] tallTokens);
        Assert.Equal(tokens, tallTokens);
        string[] shortBefore = CaptureEndnoteContinuationXml(shorter);
        string[] tallBefore = CaptureEndnoteContinuationXml(taller);
        PdfDocumentConversionResult shortResult = shorter.ConvertToPdfWithEvidence(EndnoteContinuationOptions());
        PdfDocumentConversionResult tallResult = taller.ConvertToPdfWithEvidence(EndnoteContinuationOptions());
        string[]? firstShortPages = null, firstTallPages = null;

        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] shortBytes = shortResult.ToBytes(), tallBytes = tallResult.ToBytes();
            using var shortPdf = PdfPigDocument.Open(shortBytes);
            using var tallPdf = PdfPigDocument.Open(tallBytes);
            AssertEndnoteCapacityPartitions(shortPdf, tokens, expectedCounts, bodyAdvance);
            AssertEndnoteCapacityPartitions(tallPdf, tokens, expectedCounts, bodyAdvance);
            Assert.Equal(EndnoteContinuationCompact(shortPdf.GetPage(1).Text), EndnoteContinuationCompact(tallPdf.GetPage(1).Text));
            Assert.Equal(EndnoteCapacityBaseline(shortPdf, 1, tokens[0]), EndnoteCapacityBaseline(tallPdf, 1, tokens[0]), 3);
            int firstIndex = firstCount;
            for (int page = 2; page <= expectedCounts.Length; page++) {
                // The taller measured decoration still paints before the same
                // accepted body sequence. Its advance is not a fitted tolerance.
                Assert.True(EndnoteCapacityBaseline(shortPdf, page, EndnoteContinuedSeparator) >
                    EndnoteCapacityBaseline(shortPdf, page, tokens[firstIndex]));
                Assert.True(EndnoteCapacityBaseline(tallPdf, page, EndnoteContinuedSeparator) >
                    EndnoteCapacityBaseline(tallPdf, page, tokens[firstIndex]));
                Assert.Equal(tallSeparatorAdvance - shortSeparatorAdvance,
                    EndnoteCapacityBaseline(shortPdf, page, tokens[firstIndex]) -
                    EndnoteCapacityBaseline(tallPdf, page, tokens[firstIndex]), 3);
                firstIndex += expectedCounts[page - 1];
            }
            AssertEndnoteContinuationDestinations(shorter, shortBytes, shortPdf, new[] { tokens[0] });
            AssertEndnoteContinuationDestinations(taller, tallBytes, tallPdf, new[] { tokens[0] });
            string[] shortPages = shortPdf.GetPages().Select(page => EndnoteContinuationCompact(page.Text)).ToArray();
            string[] tallPages = tallPdf.GetPages().Select(page => EndnoteContinuationCompact(page.Text)).ToArray();
            if (firstShortPages == null) { firstShortPages = shortPages; firstTallPages = tallPages; }
            else { Assert.Equal(firstShortPages, shortPages); Assert.Equal(firstTallPages, tallPages); }
            Assert.Equal(shortBefore, CaptureEndnoteContinuationXml(shorter));
            Assert.Equal(tallBefore, CaptureEndnoteContinuationXml(taller));
        }
        AssertNoEndnoteContinuationApproximation(shortResult);
        AssertNoEndnoteContinuationApproximation(tallResult);
        Assert.Empty(shorter.ValidateDocument());
        Assert.Empty(taller.ValidateDocument());
    }

    private static WordDocument CreateEndnoteCapacityDocument(int continuedSeparatorAdvance, int paragraphCount, out string[] tokens) {
        using WordDocument source = CreateEndnoteContinuationDocument(15);
        WordSection section = source.Sections[0];
        section.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(400D);
        section.PageSettings.Height = global::OfficeIMO.OfficeLength.Points(500D);
        section.Margins.Top = section.Margins.Bottom = global::OfficeIMO.OfficeLength.Points(40D);
        section.Margins.Left = section.Margins.Right = global::OfficeIMO.OfficeLength.Points(40D);
        WordParagraph reference = source.AddParagraph("CAPACITYBODYREFERENCE");
        tokens = AddEndnoteContinuationBody(reference, "ENCAPACITY", paragraphCount);
        reference._paragraph.ParagraphProperties = (W.ParagraphProperties)EndnoteContinuationParagraph(string.Empty).ParagraphProperties!.CloneNode(true);
        SetEndnoteCapacityAdvance(reference._paragraph, 20);
        foreach (W.Run run in reference._paragraph.Elements<W.Run>())
            run.RunProperties = new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" });
        ConfigureEndnoteContinuationStories(source);
        var main = source._wordprocessingDocument.MainDocumentPart!;
        W.Endnotes notes = main.EndnotesPart!.Endnotes!;
        foreach (W.Endnote note in notes.Elements<W.Endnote>().Where(WordNoteResolver.IsOrdinary))
            foreach (W.Paragraph paragraph in note.Elements<W.Paragraph>()) SetEndnoteCapacityAdvance(paragraph, 20);
        SetEndnoteCapacityAdvance(notes.Elements<W.Endnote>().Single(note => note.Type?.Value == W.FootnoteEndnoteValues.Separator)
            .Elements<W.Paragraph>().Single(), 20);
        SetEndnoteCapacityAdvance(notes.Elements<W.Endnote>().Single(note => note.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator)
            .Elements<W.Paragraph>().Single(), continuedSeparatorAdvance);
        W.Endnote notice = notes.Elements<W.Endnote>().Single(note => note.Type?.Value == W.FootnoteEndnoteValues.ContinuationNotice);
        W.EndnoteDocumentWideProperties registration = main.DocumentSettingsPart!.Settings!.GetFirstChild<W.EndnoteDocumentWideProperties>()!;
        foreach (W.EndnoteSpecialReference stored in registration.Elements<W.EndnoteSpecialReference>()
            .Where(item => item.Id?.Value == notice.Id?.Value).ToArray()) stored.Remove();
        // Keep the stored notice: this is an unselected story, not a registered
        // empty notice whose measurable paragraph must still reserve space.
        Assert.Contains(notes.Elements<W.Endnote>(), item => item.Type?.Value == W.FootnoteEndnoteValues.ContinuationNotice);
        Assert.DoesNotContain(registration.Elements<W.EndnoteSpecialReference>(), item => item.Id?.Value == notice.Id?.Value);
        Assert.Empty(source.ValidateDocument());
        WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(loaded.ValidateDocument());
        return loaded;
    }

    private static void SetEndnoteCapacityAdvance(W.Paragraph paragraph, int points) {
        W.SpacingBetweenLines spacing = paragraph.ParagraphProperties!.SpacingBetweenLines!;
        spacing.Before = "0";
        spacing.After = "0";
        spacing.Line = (points * 20).ToString(CultureInfo.InvariantCulture);
        spacing.LineRule = W.LineSpacingRuleValues.Exact;
    }

    private static void AssertEndnoteCapacityPartitions(PdfPigDocument pdf, string[] tokens, int[] counts, double advance) {
        Assert.Equal(counts.Length, pdf.NumberOfPages);
        AssertEndnoteContinuationTokenOrder(pdf, tokens);
        int offset = 0;
        for (int page = 1; page <= counts.Length; page++) {
            string text = EndnoteContinuationCompact(pdf.GetPage(page).Text);
            string[] accepted = tokens.Where(token => text.Contains(token, StringComparison.Ordinal)).ToArray();
            Assert.Equal(tokens.Skip(offset).Take(counts[page - 1]).ToArray(), accepted);
            string separator = page == 1 ? EndnoteOrdinarySeparator : EndnoteContinuedSeparator;
            Assert.Single(Regex.Matches(text, Regex.Escape(separator)));
            AssertEndnoteContinuationOrder(text, separator, accepted[0]);
            Assert.DoesNotContain(page == 1 ? EndnoteContinuedSeparator : EndnoteOrdinarySeparator, text, StringComparison.Ordinal);
            Assert.DoesNotContain(EndnoteContinuationNotice, text, StringComparison.Ordinal);
            for (int line = 1; line < accepted.Length; line++)
                Assert.Equal(advance, EndnoteCapacityBaseline(pdf, page, accepted[line - 1]) - EndnoteCapacityBaseline(pdf, page, accepted[line]), 3);
            offset += counts[page - 1];
        }
        Assert.Equal(tokens.Length, offset);
    }

    private static double EndnoteCapacityBaseline(PdfPigDocument pdf, int page, string token) {
        var letters = FindRichNoteToken(pdf.GetPage(page).Letters.ToArray(), token);
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        return letters[0].StartBaseLine.Y;
    }
}
