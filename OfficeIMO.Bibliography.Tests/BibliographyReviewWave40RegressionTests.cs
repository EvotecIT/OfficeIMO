namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave40RegressionTests {
    [Fact]
    public void Recognized_CSL_native_type_cannot_retype_an_unknown_item_during_strict_reopen() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Unknown, NativeType = "article-journal" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Fact]
    public void Additional_EndNote_record_attribute_carriers_are_diagnosed_and_omitted() {
        const string source = "<xml><records><record first=\"1\"><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "@record-attributes", "<record second=\"2\" />"));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyNativeField reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.Items).NativeFields,
            field => field.Name == "@record-attributes");

        Assert.Contains("first=\"1\"", reopened.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("second=\"2\"", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV249" && diagnostic.Field == "@record-attributes");
    }

    [Theory]
    [InlineData("titles", "<title>Second title</title>", "Second title")]
    [InlineData("periodical", "<full-title>Second journal</full-title>", "Second journal")]
    [InlineData("contributors", "<authors><author>Second, Author</author></authors>", "Second, Author")]
    [InlineData("dates", "<year>2027</year>", "2027")]
    [InlineData("urls", "<related-urls><url>https://second.example</url></related-urls>", "https://second.example")]
    [InlineData("keywords", "<keyword>second-keyword</keyword>", "second-keyword")]
    public void Retained_EndNote_containers_cannot_restore_cleared_typed_owners(string container, string nested, string marker) {
        string first = FirstContainer(container);
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type>" + first + "<" + container + ">" + nested + "</" + container + "></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);
        ClearOwner(item, container);

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain(marker, permissive.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123" && diagnostic.Field == container);
    }

    private static string FirstContainer(string container) {
        switch (container) {
            case "titles": return "<titles><title>First title</title></titles>";
            case "periodical": return "<periodical><full-title>First journal</full-title></periodical>";
            case "contributors": return "<contributors><authors><author>First, Author</author></authors></contributors>";
            case "dates": return "<dates><year>2026</year></dates>";
            case "urls": return "<urls><related-urls><url>https://first.example</url></related-urls></urls>";
            case "keywords": return "<keywords><keyword>first-keyword</keyword></keywords>";
            default: throw new ArgumentOutOfRangeException(nameof(container));
        }
    }

    private static void ClearOwner(BibliographyItem item, string container) {
        switch (container) {
            case "titles": item.Title = null; item.ContainerTitle = null; item.CollectionTitle = null; break;
            case "periodical": item.ContainerTitle = null; break;
            case "contributors": item.Contributors.Clear(); break;
            case "dates": item.Dates.Clear(); break;
            case "urls": item.Url = null; break;
            case "keywords": item.Keywords.Clear(); break;
            default: throw new ArgumentOutOfRangeException(nameof(container));
        }
    }
}
