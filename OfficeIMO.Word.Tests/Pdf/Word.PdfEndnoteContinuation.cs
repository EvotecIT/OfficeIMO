using System.Globalization;
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
    private const string EndnoteOrdinarySeparator = "ENORDINARYSEPARATOR";
    private const string EndnoteContinuedSeparator = "ENCONTINUEDSEPARATOR";
    private const string EndnoteContinuationNotice = "ENCONTINUATIONNOTICE";

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    public void SaveAsPdf_EndnoteGroupUsesDistinctContinuationStoriesWithoutRepeatingContent(int compatibilityMode) {
        using WordDocument source = CreateEndnoteContinuationDocument(compatibilityMode);
        string[] first = AddEndnoteContinuationBody(source.AddParagraph("ALPHAREFERENCE"), "ENALPHA", 40);
        string[] second = AddEndnoteContinuationBody(source.AddParagraph("BETAREFERENCE"), "ENBETA", 3);
        source.AddParagraph("ENDNOTEBODYTAIL");
        ConfigureEndnoteContinuationStories(source);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        string[] before = CaptureEndnoteContinuationXml(document);
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(EndnoteContinuationOptions());
        string[]? firstSerialization = null;

        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string[] tokens = first.Concat(second).ToArray();
            AssertEndnoteContinuationTokenOrder(pdf, tokens);
            string text = EndnoteContinuationCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
            AssertEndnoteContinuationOrder(text, "ENDNOTEBODYTAIL", first[0]);
            int firstPage = EndnoteContinuationTokenPage(pdf, first[0]);
            int finalPage = EndnoteContinuationTokenPage(pdf, second[second.Length - 1]);
            Assert.True(finalPage > firstPage, "The authored long endnote must exercise an actual page continuation.");
            Assert.Single(Regex.Matches(text, EndnoteOrdinarySeparator));

            for (int pageNumber = firstPage; pageNumber <= finalPage; pageNumber++) {
                string pageText = EndnoteContinuationCompact(pdf.GetPage(pageNumber).Text);
                string[] pageTokens = tokens.Where(token => pageText.Contains(token, StringComparison.Ordinal)).ToArray();
                Assert.NotEmpty(pageTokens);
                string separator = pageNumber == firstPage ? EndnoteOrdinarySeparator : EndnoteContinuedSeparator;
                string excluded = pageNumber == firstPage ? EndnoteContinuedSeparator : EndnoteOrdinarySeparator;
                Assert.Single(Regex.Matches(pageText, separator));
                Assert.DoesNotContain(excluded, pageText, StringComparison.Ordinal);
                AssertEndnoteContinuationOrder(pageText, separator, pageTokens[0]);
                if (pageNumber < finalPage) {
                    Assert.Single(Regex.Matches(pageText, EndnoteContinuationNotice));
                    AssertEndnoteContinuationOrder(pageText, pageTokens[pageTokens.Length - 1], EndnoteContinuationNotice);
                } else Assert.DoesNotContain(EndnoteContinuationNotice, pageText, StringComparison.Ordinal);
            }

            Assert.Equal(finalPage - firstPage, Regex.Matches(text, EndnoteContinuedSeparator).Count);
            Assert.Equal(finalPage - firstPage, Regex.Matches(text, EndnoteContinuationNotice).Count);
            AssertEndnoteContinuationDestinations(document, bytes, pdf, new[] { first[0], second[0] });
            string[] pages = pdf.GetPages().Select(page => EndnoteContinuationCompact(page.Text)).ToArray();
            if (firstSerialization == null) firstSerialization = pages;
            else Assert.Equal(firstSerialization, pages);
        }

        AssertNoEndnoteContinuationApproximation(result);
        Assert.Equal(before, CaptureEndnoteContinuationXml(document));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    public void SaveAsPdf_ShortEndnotesShareOneSeparatorWithoutAContinuationNotice(int compatibilityMode) {
        using WordDocument source = CreateEndnoteContinuationDocument(compatibilityMode);
        string[] first = AddEndnoteContinuationBody(source.AddParagraph("SHORTALPHAREFERENCE"), "ENSHORTALPHA", 1);
        string[] second = AddEndnoteContinuationBody(source.AddParagraph("SHORTBETAREFERENCE"), "ENSHORTBETA", 1);
        ConfigureEndnoteContinuationStories(source);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string[] before = CaptureEndnoteContinuationXml(document);
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(EndnoteContinuationOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = EndnoteContinuationCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
            AssertEndnoteContinuationTokenOrder(pdf, first.Concat(second).ToArray());
            Assert.Equal(EndnoteContinuationTokenPage(pdf, first[0]), EndnoteContinuationTokenPage(pdf, second[0]));
            Assert.Single(Regex.Matches(text, EndnoteOrdinarySeparator));
            Assert.DoesNotContain(EndnoteContinuedSeparator, text, StringComparison.Ordinal);
            Assert.DoesNotContain(EndnoteContinuationNotice, text, StringComparison.Ordinal);
            AssertEndnoteContinuationOrder(text, EndnoteOrdinarySeparator, first[0]);
            AssertEndnoteContinuationDestinations(document, bytes, pdf, new[] { first[0], second[0] });
        }

        AssertNoEndnoteContinuationApproximation(result);
        Assert.Equal(before, CaptureEndnoteContinuationXml(document));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SaveAsPdf_UnsupportedEndnoteGroupRetainsVisibleTextAndReportsContinuationLoss() {
        using WordDocument source = CreateEndnoteContinuationDocument(15);
        string[] first = AddEndnoteContinuationBody(source.AddParagraph("FALLBACKALPHAREFERENCE"), "ENFALLBACKALPHA", 1);
        string[] second = AddEndnoteContinuationBody(source.AddParagraph("FALLBACKBETAREFERENCE"), "ENFALLBACKBETA", 1);
        W.Endnote unsupported = source._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!
            .Elements<W.Endnote>().Last(note => note.Id?.Value > 0);
        unsupported.Elements<W.Paragraph>().Single().Append(
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" PRIVATE ENINSTRUCTIONSENTINEL ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("ENCACHEDDISPLAY")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        ConfigureEndnoteContinuationStories(source);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        string[] before = CaptureEndnoteContinuationXml(document);
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(EndnoteContinuationOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            AssertEndnoteContinuationTokenOrder(pdf, first.Concat(second).Concat(new[] { "ENCACHEDDISPLAY" }).ToArray());
            string text = EndnoteContinuationCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
            Assert.DoesNotContain("ENINSTRUCTIONSENTINEL", text, StringComparison.Ordinal);
            Assert.DoesNotContain("PRIVATE", text, StringComparison.Ordinal);
            Assert.Single(Regex.Matches(text, EndnoteOrdinarySeparator));
            Assert.DoesNotContain(EndnoteContinuedSeparator, text, StringComparison.Ordinal);
            Assert.DoesNotContain(EndnoteContinuationNotice, text, StringComparison.Ordinal);
            AssertEndnoteContinuationDestinations(document, bytes, pdf, new[] { first[0], second[0] });
        }

        PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "native-endnote-continuation-approximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Contains(result.Warnings, item => item.Code == "native-endnote-content-approximated");
        Assert.True(result.HasLoss);
        Assert.True(result.Report.HasLoss);
        Assert.Equal(PdfConversionFidelityStatus.Degraded, result.Report.FidelityStatus);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal(before, CaptureEndnoteContinuationXml(document));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_ContinuedEndnotesRetainSectionOrDocumentPlacement(bool sectionEnd, bool continuous) {
        using WordDocument source = CreateEndnoteContinuationDocument(15);
        source.AddEndnoteProperties(position: sectionEnd ? WordEndnotePosition.SectionEnd : WordEndnotePosition.DocumentEnd);
        string[] first = AddEndnoteContinuationBody(source.AddParagraph("SECTIONALPHAREFERENCE"), "ENSECTIONALPHA", 40);
        WordSection last = continuous ? source.AddSection(WordSectionBreakType.Continuous) : source.AddSection();
        last.AddParagraph("FOLLOWINGSECTIONBODY");
        string[] second = AddEndnoteContinuationBody(last.AddParagraph("SECTIONBETAREFERENCE"), "ENSECTIONBETA", 1);
        ConfigureEndnoteContinuationStories(source);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        string[] before = CaptureEndnoteContinuationXml(document);
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(EndnoteContinuationOptions());

        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            AssertEndnoteContinuationTokenOrder(pdf, first.Concat(second).ToArray());
            Assert.True(EndnoteContinuationTokenPage(pdf, first[first.Length - 1]) > EndnoteContinuationTokenPage(pdf, first[0]),
                "The first note must continue across pages before checking its section-placement boundary.");
            string text = EndnoteContinuationCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
            if (sectionEnd) AssertEndnoteContinuationOrder(text, first[first.Length - 1], "FOLLOWINGSECTIONBODY");
            else AssertEndnoteContinuationOrder(text, "FOLLOWINGSECTIONBODY", first[0]);
            AssertEndnoteContinuationOrder(text, "FOLLOWINGSECTIONBODY", second[0]);
            AssertEndnoteContinuationDestinations(document, bytes, pdf, new[] { first[0], second[0] });
        }

        AssertNoEndnoteContinuationApproximation(result);
        Assert.Equal(before, CaptureEndnoteContinuationXml(document));
        Assert.Empty(document.ValidateDocument());
    }

    private static WordDocument CreateEndnoteContinuationDocument(int compatibilityMode) {
        WordDocument document = CreateFootnoteDiagnosticDocument();
        document.CompatibilitySettings.CompatibilityMode = compatibilityMode == 14
            ? WordCompatibilityMode.Word2010 : WordCompatibilityMode.Word2013;
        return document;
    }

    private static string[] AddEndnoteContinuationBody(WordParagraph reference, string prefix, int paragraphCount) {
        reference.AddEndNote("INITIALNOTETEXT");
        W.EndnoteReference marker = reference._paragraph.Descendants<W.EndnoteReference>().Last();
        W.Endnote note = reference._document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!
            .Elements<W.Endnote>().Single(item => item.Id?.Value == marker.Id!.Value);
        note.RemoveAllChildren();
        string[] tokens = Enumerable.Range(0, paragraphCount).Select(index => prefix + index.ToString("000", CultureInfo.InvariantCulture) + "END").ToArray();
        for (int index = 0; index < tokens.Length; index++) {
            W.Paragraph paragraph = EndnoteContinuationParagraph(tokens[index], spacingAfter: "80");
            if (index == 0) paragraph.InsertAfter(new W.Run(new W.EndnoteReferenceMark()), paragraph.ParagraphProperties!);
            note.Append(paragraph);
        }
        return tokens;
    }

    private static void ConfigureEndnoteContinuationStories(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Endnotes notes = main.EndnotesPart!.Endnotes!;
        W.EndnoteDocumentWideProperties properties = main.DocumentSettingsPart!.Settings!
            .GetFirstChild<W.EndnoteDocumentWideProperties>()!;
        if (properties == null) {
            properties = new W.EndnoteDocumentWideProperties();
            main.DocumentSettingsPart.Settings.AddChild(properties, true);
        }
        properties.RemoveAllChildren<W.EndnoteSpecialReference>();
        foreach (W.Endnote special in notes.Elements<W.Endnote>().Where(note => note.Type?.Value is { } type && type != W.FootnoteEndnoteValues.Normal).ToArray())
            special.Remove();
        foreach (var story in new[] {
            (Id: -7, Type: W.FootnoteEndnoteValues.Separator, Token: EndnoteOrdinarySeparator),
            (Id: -4, Type: W.FootnoteEndnoteValues.ContinuationSeparator, Token: EndnoteContinuedSeparator),
            (Id: -13, Type: W.FootnoteEndnoteValues.ContinuationNotice, Token: EndnoteContinuationNotice)
        }) {
            notes.PrependChild(new W.Endnote(EndnoteContinuationParagraph(story.Token)) { Id = story.Id, Type = story.Type });
            properties.Append(new W.EndnoteSpecialReference { Id = story.Id });
        }
    }

    private static W.Paragraph EndnoteContinuationParagraph(string text, string spacingAfter = "0") => new(
        new W.ParagraphProperties(new W.KeepNext { Val = false }, new W.KeepLines { Val = false },
            new W.PageBreakBefore { Val = false }, new W.WidowControl { Val = false },
            new W.SpacingBetweenLines { Before = "0", After = spacingAfter, Line = "240", LineRule = W.LineSpacingRuleValues.Exact }),
        new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" }), new W.Text(text)));

    private static WordPdfExportOptions EndnoteContinuationOptions() => new() {
        IncludePageNumbers = false,
        PageSize = new PdfPageSize(400, 500),
        Margins = PdfPageMargins.Uniform(40),
        ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };

    private static string[] CaptureEndnoteContinuationXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return new[] { main.Document.OuterXml, main.EndnotesPart!.Endnotes!.OuterXml,
            main.StyleDefinitionsPart!.Styles!.OuterXml, main.DocumentSettingsPart!.Settings!.OuterXml };
    }

    private static void AssertNoEndnoteContinuationApproximation(PdfDocumentConversionResult result) =>
        Assert.DoesNotContain(result.Warnings, warning => warning.Code is "native-endnote-continuation-approximated" or "native-endnote-content-approximated");

    private static string EndnoteContinuationCompact(string text) => Regex.Replace(text, @"\s+", string.Empty);

    private static int EndnoteContinuationTokenPage(PdfPigDocument pdf, string token) =>
        Assert.Single(pdf.GetPages(), page => EndnoteContinuationCompact(page.Text).Contains(token, StringComparison.Ordinal)).Number;

    private static void AssertEndnoteContinuationTokenOrder(PdfPigDocument pdf, IReadOnlyList<string> tokens) {
        string text = EndnoteContinuationCompact(string.Concat(pdf.GetPages().Select(page => page.Text)));
        foreach (string token in tokens) Assert.Single(Regex.Matches(text, Regex.Escape(token)));
        for (int index = 1; index < tokens.Count; index++) AssertEndnoteContinuationOrder(text, tokens[index - 1], tokens[index]);
    }

    private static void AssertEndnoteContinuationOrder(string text, string first, string second) {
        Assert.Contains(first, text, StringComparison.Ordinal);
        Assert.Contains(second, text, StringComparison.Ordinal);
        Assert.True(text.IndexOf(first, StringComparison.Ordinal) < text.IndexOf(second, StringComparison.Ordinal),
            $"Expected {first} before {second}.");
    }

    private static void AssertEndnoteContinuationDestinations(WordDocument document, byte[] bytes, PdfPigDocument pdf,
        IReadOnlyList<string> firstNoteTokens) {
        W.EndnoteReference[] references = document._wordprocessingDocument.MainDocumentPart!.Document
            .Descendants<W.EndnoteReference>().ToArray();
        Assert.Equal(firstNoteTokens.Count, references.Length);
        PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
        for (int index = 0; index < references.Length; index++) {
            string name = "OfficeIMO.Endnote." + references[index].Id!.Value.ToString(CultureInfo.InvariantCulture);
            PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations, item => item.Name == name);
            PdfLinkAnnotation link = Assert.Single(navigation.LinkAnnotations, item => item.DestinationName == name);
            Assert.True(link.IsNamedDestinationLink);
            Assert.True(link.Width > 0D && link.Height > 0D);
            Assert.Equal(EndnoteContinuationTokenPage(pdf, firstNoteTokens[index]), destination.PageNumber);
        }
    }
}
