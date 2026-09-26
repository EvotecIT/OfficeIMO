using System.IO;
using System.Linq;
using System.Xml.Linq;
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
}
