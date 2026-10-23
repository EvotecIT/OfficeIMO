using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfPage = UglyToad.PdfPig.Content.Page;
using PdfLetter = UglyToad.PdfPig.Content.Letter;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfTabInheritanceTests {
    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    public void SaveAsPdf_DirectTabsComposeWithInheritedStyleTabs(string container) {
        using WordDocument document = WordDocument.Create();
        AddStyle(document, "InheritedTabs", null, Tab(4320, W.TabStopValues.Right, W.TabStopLeaderCharValues.Dot));
        W.Paragraph source = Paragraph("InheritedTabs", new W.Tabs(Tab(2880)), "FIRST", "42");
        AddToContainer(document, source, container);

        using var pdf = PdfPigDocument.Open(ConvertDocument(document));
        PdfPage page = Assert.Single(pdf.GetPages());
        double origin = container == "table" ? Marker(page, "CELLORIGIN").First().StartBaseLine.X : 72D;
        AssertStart(page, "FIRST", origin + 144D);
        AssertEnd(page, "42", origin + 216D);
        Assert.True(page.Letters.Count(letter => letter.Value == ".") >= 3,
            "The inherited right tab must retain its dot leader after a direct tab is added.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ChildClearSuppressesAncestorUnlessDirectTabRestoresThePosition(bool restoreDirectTab) {
        using WordDocument document = WordDocument.Create();
        AddStyle(document, "AncestorTabs", null, Tab(1440), Tab(4320, W.TabStopValues.Right));
        AddStyle(document, "ChildTabs", "AncestorTabs", Tab(1440, W.TabStopValues.Clear), Tab(2880));
        W.Tabs? direct = restoreDirectTab ? new W.Tabs(Tab(1440)) : null;
        W.Paragraph source = restoreDirectTab
            ? Paragraph("ChildTabs", direct, "FIRST", "SECOND", "42")
            : Paragraph("ChildTabs", direct, "FIRST", "SECOND");
        AddToContainer(document, source, "footnote");

        using var pdf = PdfPigDocument.Open(ConvertDocument(document));
        PdfPage page = Assert.Single(pdf.GetPages());
        AssertStart(page, "FIRST", restoreDirectTab ? 144D : 216D);
        if (restoreDirectTab) {
            AssertStart(page, "SECOND", 216D);
            AssertEnd(page, "42", 288D);
        } else {
            AssertEnd(page, "SECOND", 288D);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HigherPrecedenceTabReplacesAlignmentAndLeaderAtTheSamePosition(bool directOverride) {
        using WordDocument document = WordDocument.Create();
        AddStyle(document, "AncestorTabs", null, Tab(4320, W.TabStopValues.Right, W.TabStopLeaderCharValues.Dot));
        AddStyle(document, "ChildTabs", "AncestorTabs", directOverride
            ? Array.Empty<W.TabStop>() : new[] { Tab(4320) });
        W.Tabs? direct = directOverride ? new W.Tabs(Tab(4320)) : null;
        AddToContainer(document, Paragraph("ChildTabs", direct, "TARGET"), "header");

        using var pdf = PdfPigDocument.Open(ConvertDocument(document));
        PdfPage page = Assert.Single(pdf.GetPages());
        AssertStart(page, "TARGET", 288D);
        Assert.DoesNotContain(page.Letters, letter => letter.Value == ".");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_DocumentDefaultTabsParticipateInCompositionAndCanBeCleared(bool clearDefaultTab) {
        using WordDocument document = WordDocument.Create();
        W.Styles styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.DocDefaults defaults = styles.DocDefaults ?? styles.PrependChild(new W.DocDefaults());
        W.ParagraphPropertiesDefault paragraphDefaults = defaults.ParagraphPropertiesDefault ??
            defaults.AppendChild(new W.ParagraphPropertiesDefault());
        W.ParagraphPropertiesBaseStyle properties = paragraphDefaults.ParagraphPropertiesBaseStyle ??
            paragraphDefaults.AppendChild(new W.ParagraphPropertiesBaseStyle());
        properties.AddChild(new W.Tabs(Tab(1440), Tab(4320, W.TabStopValues.Right)), true);
        var direct = new W.Tabs();
        if (clearDefaultTab) direct.Append(Tab(1440, W.TabStopValues.Clear));
        direct.Append(Tab(2880));
        W.Paragraph source = clearDefaultTab
            ? Paragraph(null, direct, "FIRST", "SECOND")
            : Paragraph(null, direct, "FIRST", "SECOND", "42");
        AddToContainer(document, source, "body");

        using var pdf = PdfPigDocument.Open(ConvertDocument(document));
        PdfPage page = Assert.Single(pdf.GetPages());
        AssertStart(page, "FIRST", clearDefaultTab ? 216D : 144D);
        if (clearDefaultTab) {
            AssertEnd(page, "SECOND", 288D);
        } else {
            AssertStart(page, "SECOND", 216D);
            AssertEnd(page, "42", 288D);
        }
    }

    [Fact]
    public void SaveAsPdf_ClearedTabDoesNotShiftAnExplicitLeaderToALaterDefaultTab() {
        using WordDocument document = WordDocument.Create();
        AddStyle(document, "InheritedTabs", null, Tab(1440));
        var direct = new W.Tabs(Tab(1440, W.TabStopValues.Clear), Tab(2880),
            Tab(4320, W.TabStopValues.Right, W.TabStopLeaderCharValues.Dot));
        AddToContainer(document, Paragraph("InheritedTabs", direct, "FIRST", "42", "THIRD"), "body");

        using var pdf = PdfPigDocument.Open(ConvertDocument(document));
        PdfPage page = Assert.Single(pdf.GetPages());
        AssertStart(page, "FIRST", 216D);
        AssertEnd(page, "42", 288D);
        AssertStart(page, "THIRD", 324D);
        Assert.Contains(page.Letters, letter => letter.Value == ".");
        Assert.DoesNotContain(page.Letters, letter => letter.Value == "." && letter.StartBaseLine.X > 288D);
    }

    private static W.TabStop Tab(int position, W.TabStopValues? alignment = null,
        W.TabStopLeaderCharValues? leader = null) => new W.TabStop {
        Position = position,
        Val = alignment ?? W.TabStopValues.Left,
        Leader = leader ?? W.TabStopLeaderCharValues.None
    };

    private static void AddStyle(WordDocument document, string id, string? basedOn, params W.TabStop[] tabs) {
        var style = new W.Style(new W.StyleName { Val = id }) {
            Type = W.StyleValues.Paragraph, StyleId = id, CustomStyle = true
        };
        if (basedOn != null) style.Append(new W.BasedOn { Val = basedOn });
        if (tabs.Length > 0) style.Append(new W.StyleParagraphProperties(new W.Tabs(tabs)));
        document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
    }

    private static W.Paragraph Paragraph(string? styleId, W.Tabs? tabs, params string[] fields) {
        var properties = new W.ParagraphProperties();
        if (styleId != null) properties.Append(new W.ParagraphStyleId { Val = styleId });
        if (tabs != null) properties.Append(tabs);
        properties.Append(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact
        });
        var paragraph = new W.Paragraph(properties, Run("ORIGIN"));
        foreach (string field in fields) paragraph.Append(Run("", new W.TabChar(), new W.Text(field)));
        return paragraph;
    }

    private static W.Run Run(string text, params OpenXmlElement[] following) {
        var run = new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" },
            new W.FontSize { Val = "20" }, new W.FontSizeComplexScript { Val = "20" }));
        if (text.Length > 0) run.Append(new W.Text(text));
        foreach (OpenXmlElement element in following) run.Append(element);
        return run;
    }

    private static void AddToContainer(WordDocument document, W.Paragraph source, string container) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        if (container == "body") {
            document.AddParagraph("BODY");
            main.Document.Body!.InsertBefore(source, main.Document.Body.GetFirstChild<W.SectionProperties>());
        } else if (container == "table") {
            WordTable table = document.AddTable(1, 1);
            table.Width = 5200;
            table.WidthType = WordTableWidthUnit.Dxa;
            table.LayoutMode = WordTableLayoutMode.Fixed;
            table.Rows[0].Cells[0].Width = 5200;
            table.Rows[0].Cells[0].WidthType = WordTableWidthUnit.Dxa;
            W.TableCell cell = main.Document.Body!.Descendants<W.TableCell>().Single();
            cell.RemoveAllChildren<W.Paragraph>();
            cell.Append(new W.Paragraph(Run("CELLORIGIN")), source);
        } else if (container == "footnote") {
            document.AddParagraph("REFERENCE").AddFootNote("NOTEINTRO");
            main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
                .Single(note => note.Type == null && note.Id?.Value > 0).Append(source);
        } else {
            document.AddParagraph("BODY");
            document.AddHeadersAndFooters();
            document.Header.Default!.AddParagraph("SEED");
            W.Paragraph seed = main.HeaderParts.SelectMany(part => part.Header.Elements<W.Paragraph>())
                .Single(paragraph => paragraph.InnerText == "SEED");
            seed.InsertAfterSelf(source);
            seed.Remove();
        }
    }

    private static PdfLetter[] Marker(PdfPage page, string marker) {
        PdfLetter[] letters = page.Letters.ToArray();
        var matches = new List<PdfLetter[]>();
        for (int index = 0; index <= letters.Length - marker.Length; index++) {
            PdfLetter[] candidate = letters.Skip(index).Take(marker.Length).ToArray();
            if (string.Concat(candidate.Select(letter => letter.Value)) == marker) matches.Add(candidate);
        }
        return Assert.Single(matches);
    }

    private static void AssertStart(PdfPage page, string marker, double expected) =>
        Assert.InRange(Marker(page, marker).First().StartBaseLine.X, expected - .1D, expected + .1D);

    private static void AssertEnd(PdfPage page, string marker, double expected) =>
        Assert.InRange(Marker(page, marker).Last().EndBaseLine.X, expected - .1D, expected + .1D);

    private static byte[] ConvertDocument(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        var roots = new List<OpenXmlPartRootElement> { main.Document };
        if (main.FootnotesPart?.Footnotes is { } notes) roots.Add(notes);
        if (main.StyleDefinitionsPart?.Styles is { } styles) roots.Add(styles);
        if (main.DocumentSettingsPart?.Settings is { } settings) roots.Add(settings);
        roots.AddRange(main.HeaderParts.Select(part => Assert.IsType<W.Header>(part.Header)));
        string[] before = roots.Select(root => root.OuterXml).ToArray();
        Assert.Empty(document.ValidateDocument());
        byte[] bytes = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PageSize(612, 792),
            Margins = PageMargins.Uniform(72),
            FontFamily = "Helvetica"
        }).ToBytes();
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Equal(before, roots.Select(root => root.OuterXml).ToArray());
        Assert.Empty(document.ValidateDocument());
        return bytes;
    }
}
