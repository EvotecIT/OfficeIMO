using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfReportedTabTests {
    [Theory]
    [InlineData(14, false, 90D)]
    [InlineData(14, true, 108D)]
    [InlineData(15, false, 90D)]
    [InlineData(15, true, 90D)]
    public void SaveAsPdf_HangingImplicitTabUsesTheEffectiveCompatibilityProfile(int compatibilityMode, bool noTabHangInd, double expectedX) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = (WordCompatibilityMode)compatibilityMode;
        if (noTabHangInd) document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = true;
        var main = document.OpenXmlDocument.MainDocumentPart!;
        W.Settings settings = main.DocumentSettingsPart!.Settings!;
        settings.GetFirstChild<W.DefaultTabStop>()?.Remove();
        settings.AddChild(new W.DefaultTabStop { Val = 720 }, true);
        document.AddParagraph("BODY");
        main.Document.Body!.InsertBefore(CreateImplicitParagraph("BODYTARGET"),
            main.Document.Body.GetFirstChild<W.SectionProperties>());
        document.AddParagraph("REFERENCE").AddFootNote("FIRSTNOTE");
        main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
            .Single(note => note.Type == null && note.Id?.Value > 0).Append(CreateImplicitParagraph("NOTETARGET"));
        document.AddTable(1, 1);
        W.TableCell cell = main.Document.Body!.Descendants<W.TableCell>().Single();
        cell.RemoveAllChildren<W.Paragraph>();
        cell.Append(new W.Paragraph(TabRun("CELLORIGIN")), CreateImplicitParagraph("CELLTARGET"));

        byte[] bytes = ConvertReportedTabDocument(document);
        using var pdf = PdfPigDocument.Open(bytes);
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        foreach (string marker in new[] { "BODYTARGET", "NOTETARGET" }) {
            var target = Assert.Single(words, word => word.Text == marker);
            Assert.InRange(target.Letters.First().StartBaseLine.X, expectedX - .1D, expectedX + .1D);
        }
        double cellOrigin = Assert.Single(words, word => word.Text == "CELLORIGIN").Letters.First().StartBaseLine.X;
        var cellTarget = Assert.Single(words, word => word.Text == "CELLTARGET");
        Assert.InRange(cellTarget.Letters.First().StartBaseLine.X,
            cellOrigin + expectedX - 72.1D, cellOrigin + expectedX - 71.9D);

        static W.Paragraph CreateImplicitParagraph(string marker) => new W.Paragraph(
            new W.ParagraphProperties(
                new W.SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact },
                new W.Indentation { Left = "360", Hanging = "360" }),
            TabRun("1.", new W.TabChar(), new W.Text(marker)));
    }

    [Theory]
    [InlineData("body", 220)]
    [InlineData("body", 720)]
    [InlineData("footnote", 220)]
    [InlineData("footnote", 720)]
    [InlineData("table", 220)]
    [InlineData("table", 720)]
    public void SaveAsPdf_HangingExplicitTabRetainsItsMarginCoordinateAndIsSkippedOnLaterLines(string container, int tabStopTwips) {
        using WordDocument document = WordDocument.Create();
        W.Paragraph source = new W.Paragraph(new W.ParagraphProperties(
            new W.Tabs(new W.TabStop { Position = tabStopTwips, Val = W.TabStopValues.Left }),
            new W.SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact },
            new W.Indentation { Left = "720", Hanging = "720" }),
            TabRun("1.", new W.TabChar(), new W.Text("AFTERTAB"), new W.Break(),
                new W.Text("X"), new W.TabChar(), new W.Text("LATERLINE")));
        var main = document.OpenXmlDocument.MainDocumentPart!;
        if (container == "body") {
            document.AddParagraph("BODY");
            main.Document.Body!.InsertBefore(source, main.Document.Body.GetFirstChild<W.SectionProperties>());
        } else if (container == "table") {
            document.AddTable(1, 1);
            W.TableCell cell = main.Document.Body!.Descendants<W.TableCell>().Single();
            cell.RemoveAllChildren<W.Paragraph>();
            cell.Append(new W.Paragraph(TabRun("CELLORIGIN")), source);
        } else {
            document.AddParagraph("REFERENCE").AddFootNote("FIRSTNOTE");
            main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Single(note => note.Type == null && note.Id?.Value > 0).Append(source);
        }

        byte[] bytes = ConvertReportedTabDocument(document);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var after = TabMarkerStart(page, "AFTERTAB");
        var later = TabMarkerStart(page, "LATERLINE");
        double frameOrigin = container == "table"
            ? TabMarkerStart(page, "CELLORIGIN").StartBaseLine.X : 72D;
        double authoredStop = frameOrigin + tabStopTwips / 20D;
        Assert.InRange(after.StartBaseLine.X, authoredStop - .1D, authoredStop + .1D);
        // Both declared stops are behind the second-line cursor. The default
        // 36pt interval applies from that line's 36pt-indented text frame.
        Assert.InRange(later.StartBaseLine.X, frameOrigin + 71.9D, frameOrigin + 72.1D);
        Assert.True(later.StartBaseLine.Y < after.StartBaseLine.Y);
    }

    [Theory]
    [InlineData("body", false, false)]
    [InlineData("body", true, true)]
    [InlineData("header", false, true)]
    [InlineData("footer", false, true)]
    public void SaveAsPdf_RightAndCenterTabsAlignTheWholeStyledWordField(string container, bool center, bool styled) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODY");
        W.Paragraph source;
        if (container == "body") {
            document.AddParagraph("SEED");
            source = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().Last();
        } else {
            document.AddHeadersAndFooters();
            WordHeaderFooter story = container == "header" ? document.Header.Default! : document.Footer.Default!;
            story.AddParagraph("SEED");
            var main = document.OpenXmlDocument.MainDocumentPart!;
            source = container == "header"
                ? main.HeaderParts.SelectMany(part => part.Header.Elements<W.Paragraph>()).Single(p => p.InnerText == "SEED")
                : main.FooterParts.SelectMany(part => part.Footer.Elements<W.Paragraph>()).Single(p => p.InnerText == "SEED");
        }
        source.RemoveAllChildren<W.Run>();
        int tabStopTwips = center ? 4680 : 9000;
        source.ParagraphProperties = new W.ParagraphProperties(
            new W.Tabs(new W.TabStop { Position = tabStopTwips, Val = center ? W.TabStopValues.Center : W.TabStopValues.Right }),
            new W.SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact });
        source.Append(TabRun("Left part", new W.TabChar()));
        if (styled) {
            W.Run first = TabRun("Page");
            first.RunProperties!.Bold = new W.Bold();
            W.Run middle = TabRun(" 1 of ");
            middle.RunProperties!.Italic = new W.Italic();
            source.Append(first, middle, TabRun("3"));
        } else source.Append(TabRun("Page 1 of 3"));
        source.Append(TabRun("", new W.Break(), new W.Text("AFTERBREAK")));

        byte[] bytes = ConvertReportedTabDocument(document);
        using var pdf = PdfPigDocument.Open(bytes);
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var prefix = Assert.Single(words, word => word.Text == "Left");
        var firstWord = Assert.Single(words, word => word.Text == "Page");
        var lastWord = Assert.Single(words, word => word.Text == "3");
        var followingLine = Assert.Single(words, word => word.Text == "AFTERBREAK");
        double anchor = center
            ? (firstWord.Letters.First().StartBaseLine.X + lastWord.Letters.Last().EndBaseLine.X) / 2D
            : lastWord.Letters.Last().EndBaseLine.X;
        double authoredStop = 72D + tabStopTwips / 20D;
        Assert.InRange(anchor, authoredStop - .1D, authoredStop + .1D);
        Assert.Equal(prefix.Letters.First().StartBaseLine.Y, lastWord.Letters.First().StartBaseLine.Y, 3);
        Assert.True(followingLine.Letters.First().StartBaseLine.Y < firstWord.Letters.First().StartBaseLine.Y);
    }

    // Tight tab geometry need not produce a boundary in a reader's word heuristic.
    // These fixtures use ASCII markers; locate their painted letters directly.
    private static UglyToad.PdfPig.Content.Letter TabMarkerStart(UglyToad.PdfPig.Content.Page page, string marker) {
        var starts = Enumerable.Range(0, page.Letters.Count).Where(index =>
            string.Concat(page.Letters.Skip(index).Take(marker.Length).Select(letter => letter.Value)) == marker);
        var letters = page.Letters.Skip(Assert.Single(starts)).Take(marker.Length).ToArray();
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        return letters[0];
    }

    private static W.Run TabRun(string text, params OpenXmlElement[] following) {
        var run = new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" },
            new W.FontSize { Val = "20" }, new W.FontSizeComplexScript { Val = "20" }));
        if (text.Length > 0) run.Append(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        foreach (OpenXmlElement element in following) run.Append(element);
        return run;
    }

    private static byte[] ConvertReportedTabDocument(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        var roots = new List<OpenXmlPartRootElement> { main.Document };
        if (main.FootnotesPart?.Footnotes is { } notes) roots.Add(notes);
        if (main.StyleDefinitionsPart?.Styles is { } styles) roots.Add(styles);
        if (main.DocumentSettingsPart?.Settings is { } settings) roots.Add(settings);
        roots.AddRange(main.HeaderParts.Select(part => Assert.IsType<W.Header>(part.Header)));
        roots.AddRange(main.FooterParts.Select(part => Assert.IsType<W.Footer>(part.Footer)));
        string[] before = roots.Select(root => root.OuterXml).ToArray();
        Assert.Empty(document.ValidateDocument());
        byte[] bytes = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(612, 792),
            Margins = PdfPageMargins.Uniform(72),
            FontFamily = "Helvetica"
        }).ToBytes();
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Equal(before, roots.Select(root => root.OuterXml).ToArray());
        Assert.Empty(document.ValidateDocument());
        return bytes;
    }
}
