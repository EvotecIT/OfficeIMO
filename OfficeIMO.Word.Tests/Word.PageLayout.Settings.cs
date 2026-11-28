using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_SectionsAndRunningStoriesKeepTheirHandlesAfterSdkInterop() {
        using WordDocument document = WordDocument.Create();
        WordSection first = document.Sections[0];
        WordParagraph header = first.PageSetup.Header.AddParagraph("First");
        WordSection second = document.AddSection(WordSectionBreakType.Continuous);
        second.PageSetup.Margins.Left = global::OfficeIMO.OfficeLength.Inches(2);
        WordField footer = second.PageSetup.Footer.AddField(WordFieldType.Page);
        footer.Text = "12";

#if NET
        Assert.True(document.IsSdkDeferred);
#endif
        Assert.Equal("First", header.Text);
        Assert.Equal(WordSectionBreakType.Continuous, second.BreakType);

        _ = document.OpenXmlDocument;

        Assert.Same(first, document.Sections[0]);
        Assert.Same(second, document.Sections[1]);
        header.Text = "Updated first";
        footer.Text = "13";
        Assert.Equal("Updated first", Assert.Single(first.PageSetup.Header.Paragraphs).Text);
        Assert.Equal(global::OfficeIMO.OfficeLength.Inches(2), second.PageSetup.Margins.Left);
        Assert.Contains(document.InspectFields(), field => field.LocationKind == WordFieldLocationKind.Footer && field.ResultText == "13");
    }

    [Fact]
    public void Test_DisposedOwnedDocumentKeepsItsSections() {
        using WordDocument document = WordDocument.Create();
        List<WordSection> sections = document.Sections;
        WordSection second = document.AddSection(WordSectionBreakType.OddPage);
        document.Dispose();

        Assert.Same(sections, document.Sections);
        Assert.Same(second, document.Sections[1]);
        Assert.Equal(WordSectionBreakType.OddPage, second.BreakType);
    }

    [Fact]
    public void Test_SectionEditsReadBackFromTheActiveCompatibilityBranch() {
        byte[] bytes;
        using (WordDocument source = WordDocument.Create()) {
            var section = source.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
            var size = section.GetFirstChild<PageSize>()!;
            var choice = new AlternateContentChoice(new FootnoteProperties(), new EndnoteProperties(), size.CloneNode(true)) {
                Requires = "w14"
            };
            var fallback = new AlternateContentFallback(size.CloneNode(true));
            size.InsertBeforeSelf(new AlternateContent(choice, fallback));
            size.Remove();
            bytes = source.ToBytes();
        }

        using WordDocument document = WordDocument.Load(new MemoryStream(bytes));
        WordSection edited = document.Sections[0];
        edited.BreakType = WordSectionBreakType.Continuous;
        edited.BreakType = WordSectionBreakType.OddPage;
        Assert.Equal(WordSectionBreakType.OddPage, edited.BreakType);
        edited.PageSetup.Size.Width = OfficeLength.Inches(9);
        Assert.Equal(OfficeLength.Inches(9), edited.PageSetup.Size.Width);
        edited.FootnoteSettings.StartNumber = 5;
        edited.EndnoteSettings.StartNumber = 7;
        Assert.Equal(5, edited.FootnoteSettings.StartNumber);
        Assert.Equal(7, edited.EndnoteSettings.StartNumber);
        edited.FootnoteSettings.StartNumber = null;
        Assert.Null(edited.FootnoteSettings.StartNumber);

        var properties = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
        var alternate = Assert.Single(properties.Elements<AlternateContent>());
        var active = alternate.GetFirstChild<AlternateContentChoice>()!;
        Assert.Single(active.ChildElements, node => node.LocalName == "type");
        Assert.Equal(new[] { "pgSz" }, alternate.GetFirstChild<AlternateContentFallback>()!.ChildElements.Select(node => node.LocalName));
    }

    [Fact]
    public void Test_PageNumberInsertionPreservesSectionCompatibilityBranches() {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            var section = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
            var size = section.GetFirstChild<PageSize>()!;
            var choice = new AlternateContentChoice(size.CloneNode(true)) { Requires = "w14" };
            size.InsertBeforeSelf(new AlternateContent(choice, new AlternateContentFallback(size.CloneNode(true))));
            size.Remove();
            bytes = document.ToBytes();
        }

        using WordDocument edited = WordDocument.Load(new MemoryStream(bytes));
        edited.Sections[0].AddPageNumbering(startNumber: 2);
#if NET
        Assert.True(edited.IsSdkDeferred);
#endif
        bytes = edited.ToBytes();

        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        var properties = reopened.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
        Assert.Collection(properties.ChildElements,
            node => {
                var alternate = Assert.IsType<AlternateContent>(node);
                Assert.NotNull(alternate.GetFirstChild<AlternateContentChoice>()?.FirstChild);
                Assert.NotNull(alternate.GetFirstChild<AlternateContentFallback>()?.FirstChild);
            },
            node => Assert.IsType<PageMargin>(node),
            node => Assert.Equal(2, Assert.IsType<PageNumberType>(node).Start!.Value));
    }

    [Fact]
    public void Test_PageNumberInsertionCanUseTheActiveSectionBranch() {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            var section = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
            var margin = section.GetFirstChild<PageMargin>()!;
            var choice = new AlternateContentChoice(margin.CloneNode(true), new Columns()) { Requires = "w14" };
            var fallback = new AlternateContentFallback(margin.CloneNode(true), new Columns());
            margin.InsertBeforeSelf(new AlternateContent(choice, fallback));
            margin.Remove();
            bytes = document.ToBytes();
        }

        using WordDocument edited = WordDocument.Load(new MemoryStream(bytes));
        edited.Sections[0].AddPageNumbering(startNumber: 2);
        Assert.Equal(2, edited.Sections[0].PageNumberSettings.StartNumber);
#if NET
        Assert.True(edited.IsSdkDeferred);
#endif
        bytes = edited.ToBytes();

        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        var properties = reopened.OpenXmlDocument.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;
        var alternate = Assert.Single(properties.Elements<AlternateContent>());
        Assert.Equal(new[] { "pgMar", "pgNumType", "cols" },
            alternate.GetFirstChild<AlternateContentChoice>()!.ChildElements.Select(node => node.LocalName));
        Assert.Equal(new[] { "pgMar", "cols" },
            alternate.GetFirstChild<AlternateContentFallback>()!.ChildElements.Select(node => node.LocalName));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_PageLayoutSettingsPreserveMarginsAndOtherDocumentFlags(bool mirrorMargins, bool gutterAtTop) {
        using WordDocument document = WordDocument.Create();
        document.Settings.MirrorMargins = mirrorMargins;
        document.Settings.GutterAtTop = gutterAtTop;
        document.Settings.TrackRevisions = true;
        document.Sections[0].PageSetup.DifferentOddAndEvenPages = true;
        document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal,
            WordEndnotePosition.DocumentEnd, WordNoteNumberRestart.Continuous, startNumber: 1);
        document.Sections[0].PageSetup.Margins.Gutter = global::OfficeIMO.OfficeLength.Points(20D);
        document.AddParagraph("Margin settings");
        byte[] bytes = document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });

        // Check the MS-DOC wire contract independently of the importer.
        byte[] wordStream = ReadCompoundStream(bytes, "WordDocument");
        byte[] tableStream = ReadCompoundStream(bytes, "1Table");
        int dopOffset = BitConverter.ToInt32(wordStream, 0x192);
        int dopLength = BitConverter.ToInt32(wordStream, 0x196);
        Assert.InRange(dopOffset, 0, tableStream.Length - dopLength);
        Assert.True(dopLength >= (gutterAtTop ? 84 : 56));
        Assert.Equal(mirrorMargins, (tableStream[dopOffset + 6] & 0x20) != 0);
        if (dopLength >= 84) Assert.Equal(gutterAtTop, (tableStream[dopOffset + 83] & 0x80) != 0);

        using var source = new MemoryStream(bytes);
        using WordDocument loaded = WordDocument.Load(source);
        Assert.Equal(mirrorMargins, loaded.Settings.MirrorMargins);
        Assert.Equal(gutterAtTop, loaded.Settings.GutterAtTop);
        Assert.True(loaded.Settings.TrackRevisions);
        Assert.True(loaded.Sections[0].PageSetup.DifferentOddAndEvenPages);
        Assert.Equal(400U, checked((uint)(loaded.Sections[0].PageSetup.Margins.Gutter.GetValueOrDefault().ToTwips())));
        Assert.Equal(EndnotePositionValues.DocumentEnd,
            loaded.Sections[0].EndnoteProperties.EndnotePosition!.Val!.Value);
        Assert.Equal("Margin settings", Assert.Single(loaded.Paragraphs).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageLayout_ImplicitOnOffValuesRemainEnabledThroughNativeDoc(bool topGutter) {
        using WordDocument document = WordDocument.Create();
        if (topGutter) {
            var settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            settings.AddChild(new GutterAtTop(), true);
            Assert.True(document.Settings.GutterAtTop);
        } else {
            document.Sections[0]._sectionProperties.AddChild(new GutterOnRight(), true);
            Assert.True(document.Sections[0].RtlGutter);
        }
        document.AddParagraph("Implicit gutter");
        using var source = new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
        using WordDocument loaded = WordDocument.Load(source);
        Assert.Equal(topGutter, loaded.Settings.GutterAtTop);
        Assert.Equal(!topGutter, loaded.Sections[0].RtlGutter);
        loaded.Settings.GutterAtTop = false;
        loaded.Sections[0].RtlGutter = false;
        Assert.False(loaded.Settings.GutterAtTop);
        Assert.False(loaded.Sections[0].RtlGutter);
    }

    [Fact]
    public void PageLayout_RightGutterInsertionRetainsSectionSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        var properties = document.Sections[0]._sectionProperties;
        properties.AddChild(new DocGrid { Type = DocGridValues.Lines }, true);
        var validator = new OpenXmlValidator();
        Assert.Empty(validator.Validate(properties));
        document.Sections[0].RtlGutter = true;
        Assert.Empty(validator.Validate(properties));
    }
}
