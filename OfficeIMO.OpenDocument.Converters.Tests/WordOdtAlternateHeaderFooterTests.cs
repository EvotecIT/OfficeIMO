using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtAlternateHeaderFooterTests {
    [Fact]
    public void WordFirstAndEvenStoriesRoundTripThroughNativeOdt() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Body");
        WordSection section = word.Sections[0];
        section.AddHeadersAndFooters();
        section.DifferentFirstPage = true;
        section.DifferentOddAndEvenPages = true;
        section.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Odd header");
        section.GetOrCreateFooter(WordHeaderFooterType.Default).AddParagraph("Odd footer");
        section.GetOrCreateHeader(WordHeaderFooterType.First).AddParagraph("First header");
        section.GetOrCreateFooter(WordHeaderFooterType.First).AddParagraph("First footer");
        section.GetOrCreateHeader(WordHeaderFooterType.Even).AddParagraph("Even header");
        section.GetOrCreateFooter(WordHeaderFooterType.Even).AddParagraph("Even footer");

        OdfConversionResult<OdtDocument> conversion = word.ToOpenDocumentResult();
        OdtDocument odt = OdtDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        Assert.Equal("Odd header", Assert.Single(odt.PageLayout.Header.Paragraphs).Text);
        Assert.Equal("Odd footer", Assert.Single(odt.PageLayout.Footer.Paragraphs).Text);
        Assert.Equal("First header", Assert.Single(odt.PageLayout.FirstHeader!.Paragraphs).Text);
        Assert.Equal("First footer", Assert.Single(odt.PageLayout.FirstFooter!.Paragraphs).Text);
        Assert.Equal("Even header", Assert.Single(odt.PageLayout.LeftHeader!.Paragraphs).Text);
        Assert.Equal("Even footer", Assert.Single(odt.PageLayout.LeftFooter!.Paragraphs).Text);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "alternate-headers-footers" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        word.ToOpenDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        });

        using WordDocument reopened = WordDocument.Load(new MemoryStream(odt.ToWordDocument().ToBytes()));
        WordSection roundTrip = reopened.Sections[0];
        Assert.True(roundTrip.DifferentFirstPage);
        Assert.True(roundTrip.DifferentOddAndEvenPages);
        Assert.Equal("Odd header", Assert.Single(roundTrip.Header.Default!.Paragraphs).Text);
        Assert.Equal("First header", Assert.Single(roundTrip.Header.First!.Paragraphs).Text);
        Assert.Equal("Even header", Assert.Single(roundTrip.Header.Even!.Paragraphs).Text);
        Assert.Equal("Odd footer", Assert.Single(roundTrip.Footer.Default!.Paragraphs).Text);
        Assert.Equal("First footer", Assert.Single(roundTrip.Footer.First!.Paragraphs).Text);
        Assert.Equal("Even footer", Assert.Single(roundTrip.Footer.Even!.Paragraphs).Text);
        odt.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        }).Value.Dispose();
    }

    [Fact]
    public void EmptyFirstAndLeftVariantsSuppressDefaultStoryAfterWordConversion() {
        OdtDocument odt = OdtDocument.Create();
        odt.AddParagraph("Body");
        odt.PageLayout.Header.AddParagraph("Ordinary header");
        odt.PageLayout.EnsureFirstHeader();
        odt.PageLayout.EnsureLeftHeader();

        using WordDocument word = odt.ToWordDocument();
        WordSection section = word.Sections[0];
        Assert.True(section.DifferentFirstPage);
        Assert.True(section.DifferentOddAndEvenPages);
        Assert.Equal("Ordinary header", Assert.Single(section.Header.Default!.Paragraphs).Text);
        Assert.DoesNotContain(section.Header.First!.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
        Assert.DoesNotContain(section.Header.Even!.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
    }

    [Fact]
    public void AlternateHeaderTableIsReportedAsLossBeforeWordConversion() {
        OdtDocument odt = OdtDocument.Create();
        odt.AddParagraph("Body");
        odt.PageLayout.EnsureFirstHeader().AddParagraph("Visible first header");
        XElement firstHeader = odt.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "header-first").Single();
        firstHeader.Add(new XElement(OdfNamespaces.Table + "table",
            new XElement(OdfNamespaces.Table + "table-row",
                new XElement(OdfNamespaces.Table + "table-cell",
                    new XElement(OdfNamespaces.Text + "p", "Unmapped table text")))));
        odt.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = odt.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Equal("Visible first header", Assert.Single(word.Sections[0].Header.First!.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "header-footer-blocks" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => odt.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        }));
    }

    [Fact]
    public void ExplicitlyDisabledWordFirstPageDoesNotReplaceDefaultHeader() {
        using WordDocument authored = WordDocument.Create();
        authored.AddParagraph("Body");
        WordSection section = authored.Sections[0];
        section.AddHeadersAndFooters();
        section.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Default header");
        section.DifferentFirstPage = true;
        section.GetOrCreateHeader(WordHeaderFooterType.First).AddParagraph("Dormant first header");

        using var stream = new MemoryStream();
        byte[] packageBytes = authored.ToBytes();
        stream.Write(packageBytes, 0, packageBytes.Length);
        stream.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            SectionProperties properties = package.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>().Single();
            properties.GetFirstChild<TitlePage>()!.Val = false;
            package.MainDocumentPart.Document!.Save();
        }
        using WordDocument source = WordDocument.Load(new MemoryStream(stream.ToArray()));
        Assert.False(source.Sections[0].DifferentFirstPage);
        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Default header", Assert.Single(conversion.Value.PageLayout.Header.Paragraphs).Text);
        Assert.Null(conversion.Value.PageLayout.FirstHeader);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "alternate-headers-footers" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        source.Sections[0].DifferentFirstPage = true;
        Assert.True(source.Sections[0].DifferentFirstPage);
        using WordprocessingDocument reenabled = WordprocessingDocument.Open(new MemoryStream(source.ToBytes()), false);
        Assert.Single(reenabled.MainDocumentPart!.Document!.Body!.Descendants<TitlePage>());
    }

    [Fact]
    public void WordOddEvenSettingWithMissingEvenPartsEmitsBlankLeftStories() {
        using WordDocument authored = WordDocument.Create();
        authored.AddParagraph("Body");
        WordSection section = authored.Sections[0];
        section.AddHeadersAndFooters();
        section.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Odd header");
        section.GetOrCreateFooter(WordHeaderFooterType.Default).AddParagraph("Odd footer");
        section.DifferentOddAndEvenPages = true;

        using var stream = new MemoryStream();
        byte[] packageBytes = authored.ToBytes();
        stream.Write(packageBytes, 0, packageBytes.Length);
        stream.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            SectionProperties properties = package.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>().Single();
            foreach (HeaderReference reference in properties.Elements<HeaderReference>()
                         .Where(reference => reference.Type?.Value == HeaderFooterValues.Even).ToList()) reference.Remove();
            foreach (FooterReference reference in properties.Elements<FooterReference>()
                         .Where(reference => reference.Type?.Value == HeaderFooterValues.Even).ToList()) reference.Remove();
            package.MainDocumentPart.Document!.Save();
        }
        using WordDocument source = WordDocument.Load(new MemoryStream(stream.ToArray()));
        Assert.False(source.Sections[0].DifferentOddAndEvenPages);
        Assert.True(source.CreateInspectionSnapshot().Sections[0].DocumentOddEvenSettingEnabled);
        OdtDocument odt = source.ToOpenDocument();
        Assert.Equal("Odd header", Assert.Single(odt.PageLayout.Header.Paragraphs).Text);
        Assert.Equal("Odd footer", Assert.Single(odt.PageLayout.Footer.Paragraphs).Text);
        Assert.Empty(odt.PageLayout.LeftHeader!.Paragraphs);
        Assert.Empty(odt.PageLayout.LeftFooter!.Paragraphs);
    }

    [Fact]
    public void HiddenOdtStoriesDoNotBecomeVisibleInWord() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.Header.AddParagraph("Hidden default");
        source.PageLayout.Header.IsDisplayed = false;
        source.PageLayout.EnsureFirstHeader().AddParagraph("Hidden first");
        source.PageLayout.FirstHeader!.IsDisplayed = false;
        source.PageLayout.EnsureLeftHeader().AddParagraph("Hidden even");
        source.PageLayout.LeftHeader!.IsDisplayed = false;

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.True(word.Sections[0].DifferentFirstPage);
        Assert.True(word.Sections[0].DifferentOddAndEvenPages);
        Assert.DoesNotContain(word.Sections[0].Header.Default!.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
        Assert.DoesNotContain(word.Sections[0].Header.First!.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
        Assert.DoesNotContain(word.Sections[0].Header.Even!.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "hidden-header-footer-content" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 3);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        }));
    }

    [Fact]
    public void MissingOdtAlternateUsesItsDefaultHeaderOrFooter() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.Header.AddParagraph("Default header");
        source.PageLayout.Footer.AddParagraph("Default footer");
        source.PageLayout.EnsureFirstHeader().AddParagraph("First header");
        source.PageLayout.EnsureLeftFooter().AddParagraph("Even footer");

        using WordDocument word = WordDocument.Load(new MemoryStream(source.ToWordDocument().ToBytes()));
        WordSection section = word.Sections[0];
        Assert.True(section.DifferentFirstPage);
        Assert.True(section.DifferentOddAndEvenPages);
        Assert.Equal("First header", Assert.Single(section.Header.First!.Paragraphs).Text);
        Assert.Equal("Default footer", Assert.Single(section.Footer.First!.Paragraphs).Text);
        Assert.Equal("Default header", Assert.Single(section.Header.Even!.Paragraphs).Text);
        Assert.Equal("Even footer", Assert.Single(section.Footer.Even!.Paragraphs).Text);
        source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        }).Value.Dispose();

        OdtDocument reverse = OdtDocument.Create();
        reverse.AddParagraph("Body");
        reverse.PageLayout.Header.AddParagraph("Default header");
        reverse.PageLayout.Footer.AddParagraph("Default footer");
        reverse.PageLayout.EnsureFirstFooter().AddParagraph("First footer");
        reverse.PageLayout.EnsureLeftHeader().AddParagraph("Even header");
        using WordDocument reversed = WordDocument.Load(new MemoryStream(reverse.ToWordDocument().ToBytes()));
        Assert.Equal("Default header", Assert.Single(reversed.Sections[0].Header.First!.Paragraphs).Text);
        Assert.Equal("First footer", Assert.Single(reversed.Sections[0].Footer.First!.Paragraphs).Text);
        Assert.Equal("Even header", Assert.Single(reversed.Sections[0].Header.Even!.Paragraphs).Text);
        Assert.Equal("Default footer", Assert.Single(reversed.Sections[0].Footer.Even!.Paragraphs).Text);
    }

    [Fact]
    public void FirstPageStoriesCannotBeSavedAsOdf13() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.EnsureFirstHeader().AddParagraph("First page");

        Assert.Throws<InvalidOperationException>(() => source.ToBytes(new OdfSaveOptions {
            CompatibilityProfile = OdfCompatibilityProfile.Odf13
        }));
        Assert.NotEmpty(source.ToBytes());

        OdtDocument older = OdtDocument.Load(new MemoryStream(OdtDocument.Create().ToBytes(
            new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 })));
        older.PageLayout.EnsureFirstFooter().AddParagraph("First footer");
        Assert.Throws<InvalidOperationException>(() => older.ToBytes(new OdfSaveOptions {
            CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
        }));
        Assert.Throws<InvalidOperationException>(() => older.ToFlatXml());
        using var flatStream = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => older.SaveFlatXml(flatStream));
        Assert.Equal(0, flatStream.Length);
        string flatPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".fodt");
        Assert.Throws<InvalidOperationException>(() => older.SaveFlatXml(flatPath));
        Assert.False(File.Exists(flatPath));
    }

    [Fact]
    public void FallbackHeaderDoesNotDoubleCountSourceHyperlink() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.Header.AddParagraph().AddHyperlink("Link", "https://example.test/");
        source.PageLayout.EnsureFirstFooter().AddParagraph("First footer");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Equal("Link", target.Sections[0].Header.First!.Paragraphs.Single().Text);
        Assert.Contains(conversion.Report.ForFeature("hyperlinks"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
    }

    [Fact]
    public void AlternateHeadingKeepsItsWordHeadingStyle() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.EnsureFirstHeader();
        XElement first = source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "header-first").Single();
        first.Add(new XElement(OdfNamespaces.Text + "h",
            new XAttribute(OdfNamespaces.Text + "outline-level", "2"), "Heading"));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Equal(WordParagraphStyles.Heading2,
            target.Sections[0].Header.First!.Paragraphs.Single().Style);
    }

    [Fact]
    public void OutlineLevelTenReportsApproximationInBodyAndFirstHeader() {
        OdtDocument source = OdtDocument.Create();
        source.AddHeading("Body heading", 10);
        source.PageLayout.EnsureFirstHeader().AddHeading("Header heading", 10);

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Equal(WordParagraphStyles.Heading9,
            target.Sections[0].Header.First!.Paragraphs.Single().Style);
        Assert.Contains(conversion.Report.ForFeature("heading-levels"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 2);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordHeadingInFirstHeaderKeepsItsOdtOutlineLevel() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body");
        WordSection section = source.Sections[0];
        section.AddHeadersAndFooters();
        section.DifferentFirstPage = true;
        WordParagraph heading = section.GetOrCreateHeader(WordHeaderFooterType.First).AddParagraph("Heading");
        heading.Style = WordParagraphStyles.Heading2;

        OdtDocument target = source.ToOpenDocument();
        Assert.Equal(2, target.PageLayout.FirstHeader!.Paragraphs.Single().HeadingLevel);
        Assert.Contains(target.Package.GetXml("styles.xml").Descendants(OdfNamespaces.Text + "h"),
            element => (string?)element.Attribute(OdfNamespaces.Text + "outline-level") == "2");
    }

    [Fact]
    public void FallbackHeaderDoesNotDoubleCountUnsupportedSourceNote() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.Header.AddParagraph().AddFootnote("Header note");
        source.PageLayout.EnsureFirstFooter().AddParagraph("First footer");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("note-headers-footers"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void BasicFieldInFirstHeaderIsMappedWithoutFalseSourceLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body");
        source.PageLayout.EnsureFirstHeader().AddParagraph().AddField(OdtFieldKind.PageNumber, "7");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.DoesNotContain(conversion.Report.ForFeature("source-text-fields"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(conversion.Report.ForFeature("fields"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
    }

    [Fact]
    public void EmptyLaterDefaultHeaderIsReportedAsLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body");
        source.Sections[0].AddHeadersAndFooters();
        source.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("First section");
        WordSection later = source.AddSection();
        later.AddParagraph("Later body");
        later.AddHeadersAndFooters();
        later.GetOrCreateHeader(WordHeaderFooterType.Default);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.ForFeature("section-headers-footers"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count >= 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
        }));
    }

    [Fact]
    public void LaterFirstPageSettingWithoutExplicitPartsIsReportedAsLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("First body");
        source.Sections[0].AddHeadersAndFooters();
        source.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Odd header");
        WordSection later = source.AddSection();
        later.AddParagraph("Second body");
        later.DifferentFirstPage = true;

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.ForFeature("alternate-headers-footers"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }));
    }

    [Fact]
    public void EvenPageNumberRestartIsReportedBeforeAlternateStoryConversion() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body");
        WordSection section = source.Sections[0];
        section.AddHeadersAndFooters();
        section.DifferentOddAndEvenPages = true;
        section.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Odd header");
        section.GetOrCreateHeader(WordHeaderFooterType.Even).AddParagraph("Even header");
        section.AddPageNumbering(startNumber: 2);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.ForFeature("page-numbering"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }));
    }

    [Fact]
    public void LaterSectionRestartAtOneIsReportedAsLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("First body");
        WordSection later = source.AddSection();
        later.AddParagraph("Second body");
        later.AddPageNumbering(startNumber: 1);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.ForFeature("page-numbering"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }));
    }

    [Fact]
    public void InheritedEvenHeaderInLaterSectionIsNotAnotherLostStory() {
        using WordDocument authored = WordDocument.Create();
        authored.AddParagraph("Body");
        authored.Sections[0].AddHeadersAndFooters();
        authored.Sections[0].DifferentOddAndEvenPages = true;
        authored.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Even).AddParagraph("Even header");
        WordSection later = authored.AddSection();
        later.AddParagraph("Later body");

        using var stream = new MemoryStream();
        byte[] bytes = authored.ToBytes();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            SectionProperties[] sections = package.MainDocumentPart!.Document!.Body!
                .Descendants<SectionProperties>().ToArray();
            Assert.Equal(2, sections.Length);
            foreach (HeaderReference reference in sections[1].Elements<HeaderReference>().ToArray())
                reference.Remove();
            foreach (FooterReference reference in sections[1].Elements<FooterReference>().ToArray())
                reference.Remove();
            package.MainDocumentPart.Document.Save();
        }

        using WordDocument source = WordDocument.Load(new MemoryStream(stream.ToArray()));
        var inherited = source.CreateInspectionSnapshot().Sections[1];
        Assert.NotNull(inherited.EvenHeader);
        Assert.False(inherited.HasExplicitEvenHeader);
        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.ForFeature("alternate-headers-footers"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        var first = source.CreateInspectionSnapshot().Sections[0];
        int firstStoryBlocks = new[] { first.DefaultHeader, first.DefaultFooter, first.FirstHeader,
                first.FirstFooter, first.EvenHeader, first.EvenFooter }
            .Where(part => part != null).Sum(part => part!.Elements.Count);
        OdfConversionResult<OdtDocument> omitted = source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { IncludeHeadersAndFooters = false });
        Assert.Equal(firstStoryBlocks, Assert.Single(omitted.Report.ForFeature("headers-footers")).Count);
    }
}
