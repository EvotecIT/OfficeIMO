using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableOfContentsReadbackTests {
    private static readonly (string Title, int Level)[] NestedSections = {
        ("Overview", 1),
        ("Scope", 2),
        ("Audience", 2),
        ("Findings", 1),
        ("Next steps", 1)
    };

    [Theory]
    [InlineData(PdfComplianceProfile.PdfUa1)]
    [InlineData(PdfComplianceProfile.PdfUa2)]
    public void TaggedContents_UseTocStructureWithArtifactLeaders(PdfComplianceProfile profile) {
        byte[] bytes = CreateContents(Options(profile), NestedSections, startOnNewPage: true);

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).TaggedContent);
        var byObject = tagged.StructureElements.ToDictionary(static element => element.ObjectNumber);
        IEnumerable<PdfStructureElementInfo> Children(PdfStructureElementInfo element, string type) =>
            element.ChildElementObjectNumbers.Select(number => byObject[number]).Where(child => child.StructureType == type);

        PdfStructureElementInfo root = Assert.Single(tagged.StructureElements, element => element.StructureType == "TOC" &&
            byObject[element.ParentObjectNumber!.Value].StructureType != "TOCI");
        PdfStructureElementInfo[] items = Children(root, "TOCI").ToArray();
        Assert.Equal(3, items.Length);
        PdfStructureElementInfo nested = Assert.Single(Children(items[0], "TOC"));
        Assert.Equal(2, Children(nested, "TOCI").Count());
        Assert.Empty(Children(items[1], "TOC"));
        foreach (PdfStructureElementInfo item in tagged.StructureElements.Where(static element => element.StructureType == "TOCI")) {
            PdfStructureElementInfo reference = Assert.Single(Children(item, "Reference"));
            PdfStructureElementInfo link = Assert.Single(Children(reference, "Link"));
            Assert.True(link.ObjectReferenceCount > 0);
        }

        PdfTextSpan[] leaders = PdfDocument.Load(bytes, new PdfLoadOptions { IncludeArtifactText = true }).Read().Pages[0].TextBlocks
            .SelectMany(static block => block.Spans)
            .Where(static span => span.Text.Trim().Length >= 3 && span.Text.Trim().All(static ch => ch == '.'))
            .ToArray();
        Assert.NotEmpty(leaders);
        Assert.All(leaders, static span => Assert.True(span.IsArtifactContent));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NestedContents_ReadAsOneListOfLinkedEntries(bool tagged) {
        byte[] bytes = CreateContents(Options(tagged), NestedSections, startOnNewPage: true);

        PdfDocumentReadResult read = PdfDocument.Load(bytes).Read();
        PdfLogicalPage contents = read.Pages[0];
        PdfLogicalListItem[] entries = contents.ListItems.Where(static item => item.IsContentsEntry).ToArray();

        Assert.Empty(contents.Tables);
        Assert.DoesNotContain(contents.Paragraphs, static paragraph => paragraph.Text.Contains("Overview", StringComparison.Ordinal));
        Assert.Equal(NestedSections.Select(static section => section.Title), entries.Select(static entry => entry.Text));
        Assert.Equal(new[] { 1, 2, 2, 1, 1 }, entries.Select(static entry => entry.Level));
        Assert.Equal(new[] { "2", "3", "4", "5", "6" }, entries.Select(static entry => entry.TargetPageLabel));
        Assert.All(entries, static entry => Assert.StartsWith("section-", entry.Link?.DestinationName, StringComparison.Ordinal));

        string[] markdown = contents.ExtractMarkdown().Split('\n')
            .Select(static line => line.TrimEnd('\r'))
            .Where(static line => line.TrimStart().StartsWith("- ", StringComparison.Ordinal)).ToArray();
        Assert.Equal(5, markdown.Length);
        Assert.Matches(@"^- \[Overview\]\(#section-[^)]+\) — 2$", markdown[0]);
        Assert.Matches(@"^  - \[Scope\]\(#section-[^)]+\) — 3$", markdown[1]);
        Assert.Matches(@"^- \[Next steps\]\(#section-[^)]+\) — 6$", markdown[4]);
    }

    [Fact]
    public void TaggedAndUntaggedContents_ReadToTheSameEntries() {
        static (string, int, string?, string?)[] Read(bool tagged) => PdfDocument.Load(CreateContents(Options(tagged), NestedSections, startOnNewPage: true))
            .Read().Pages[0].ListItems.Where(static item => item.IsContentsEntry)
            .Select(static item => (item.Text, item.Level, item.TargetPageLabel, item.Link?.DestinationName)).ToArray();

        Assert.Equal(Read(tagged: true), Read(tagged: false));
    }

    [Fact]
    public void ContentsEntries_KeepTheirPageNumbersInWordHtmlAndRtf() {
        byte[] bytes = CreateContents(Options(tagged: true), NestedSections, startOnNewPage: true);
        string expected = "Scope — 3";

        using (OfficeIMO.Word.WordDocument word = PdfDocument.Load(bytes).ToWord(new PdfWordExportOptions()))
        using (var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(new System.IO.MemoryStream(word.ToBytes()), false)) {
            Assert.Contains(docx.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>(),
                paragraph => paragraph.ParagraphProperties?.NumberingProperties is not null && paragraph.InnerText == expected);
        }
        string html = System.Net.WebUtility.HtmlDecode(PdfDocument.Load(bytes).ToHtml(new PdfHtmlExportOptions()));
        Assert.Contains("<li>" + expected + "</li>", html, StringComparison.Ordinal);
        OfficeIMO.Rtf.RtfDocument rtf = PdfDocument.Load(bytes).Read().ToRtf(new PdfRtfExportOptions());
        Assert.Contains(rtf.Paragraphs, paragraph => paragraph.ToPlainText() == expected);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ContentsSpanningTwoPages_ReadAsListsOnBothPages(bool tagged) {
        (string, int)[] sections = Enumerable.Range(1, 70).Select(static index => ("Topic " + index, 1)).ToArray();
        byte[] bytes = CreateContents(Options(tagged), sections, startOnNewPage: false);

        PdfDocumentReadResult read = PdfDocument.Load(bytes).Read();
        PdfLogicalListItem[] first = read.Pages[0].ListItems.Where(static item => item.IsContentsEntry).ToArray();
        PdfLogicalListItem[] second = read.Pages[1].ListItems.Where(static item => item.IsContentsEntry).ToArray();

        Assert.NotEmpty(first);
        Assert.NotEmpty(second);
        Assert.Empty(read.Pages[0].Tables);
        Assert.Empty(read.Pages[1].Tables);
        Assert.Equal(sections.Select(static section => section.Item1), first.Concat(second).Select(static entry => entry.Text));
        Assert.All(first.Concat(second), static entry => Assert.NotNull(entry.TargetPageLabel));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TwoColumnTableWithNumbers_StaysATable(bool tagged) {
        PdfDocument document = PdfDocument.Create(Options(tagged)).Meta(title: "Regional counts");
        document.AddTable(new[] {
            new[] { "Region", "Sites" },
            new[] { "North", "2" },
            new[] { "South", "3" },
            new[] { "East", "4" },
            new[] { "West", "5" }
        });
        // Enough pages that every count is also a valid page number.
        for (int index = 2; index <= 5; index++) document.PageBreak().AddParagraph("Page " + index);
        byte[] bytes = document.ToBytes();

        PdfLogicalPage page = PdfDocument.Load(bytes).Read().Pages[0];

        Assert.DoesNotContain(page.ListItems, static item => item.IsContentsEntry);
        PdfLogicalTable table = Assert.Single(page.Tables);
        Assert.Equal(new[] { "West", "5" }, table.Rows[table.Rows.Count - 1]);
    }

    [Theory]
    [InlineData(true, "2.50", "3.10", "3.40", "4.25")]
    [InlineData(false, "2.50", "3.10", "3.40", "4.25")]
    [InlineData(false, "2", "3", "3", "4")]
    public void DotLeaderPriceList_IsNotContents(bool tagged, string espresso, string latte, string mocha, string cocoa) {
        PdfDocument document = PdfDocument.Create(Options(tagged)).Meta(title: "Menu").H1("Menu");
        foreach ((string item, string price) in new[] { ("Espresso", espresso), ("Latte", latte), ("Mocha", mocha), ("Cocoa", cocoa) }) {
            var style = new PdfParagraphStyle();
            style.AddTabStop(global::OfficeIMO.OfficeLength.Points(1D), PdfTabAlignment.Right, PdfTabLeaderStyle.Dots).IsFrameRightRelative = true;
            document.AddParagraph(paragraph => paragraph.AddRuns(new[] {
                PdfTextRun.Normal(item), PdfTextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right), PdfTextRun.Normal(price)
            }), style: style);
        }
        for (int index = 2; index <= 5; index++) document.PageBreak().AddParagraph("Page " + index);

        PdfLogicalPage first = PdfDocument.Load(document.ToBytes()).Read().Pages[0];

        Assert.DoesNotContain(first.ListItems, static item => item.IsContentsEntry);
        Assert.Contains(first.Tables, table => table.Rows.Any(row => row.Contains(espresso) || row.Any(cell => cell.Contains("Espresso"))));
    }

    private static PdfCreateOptions Options(bool tagged) => Options(tagged ? PdfComplianceProfile.PdfUa1 : PdfComplianceProfile.None);

    // The bundled font satisfies the PDF/UA Unicode-mapping gate; untagged output uses it too, so only tagging differs.
    private static PdfCreateOptions Options(PdfComplianceProfile profile) {
        string? font = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(font);
        PdfCreateOptions options = new PdfCreateOptions().UseFontFamily("Carlito", font!);
        return profile == PdfComplianceProfile.None ? options : options.UsePdfUa(profile, "en-US");
    }

    private static byte[] CreateContents(PdfCreateOptions options, IReadOnlyList<(string Title, int Level)> sections, bool startOnNewPage) {
        PdfDocument document = PdfDocument.Create(options)
            .Meta(title: "Contents readback")
            .TableOfContents(new PdfTableOfContentsOptions { MaximumLevel = 3 });
        foreach ((string title, int level) in sections) {
            document.Section(title, section => section.AddParagraph(title + " body."),
                new PdfSectionOptions { Level = level, StartOnNewPage = startOnNewPage });
        }
        return document.ToBytes();
    }
}
