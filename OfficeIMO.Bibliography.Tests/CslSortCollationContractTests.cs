using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslSortCollationContractTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PunctuationDoesNotMoveNamesAheadOfTheirAlphabeticPosition(bool macro, bool html) {
        const string json = """
            [{"id":"roe","type":"book","author":[{"family":"Roe"}]},
             {"id":"flinders","type":"book","author":[{"family":"[F]linders"}]},
             {"id":"doe","type":"book","author":[{"family":"Doe"}]}]
            """;
        CslRenderResult result = Render(json, macro ? "macro=\"names\"" : "variable=\"author\"", "<names variable=\"author\"><name/></names>", html: html);
        Assert.Equal(new[] { "doe", "flinders", "roe" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Equal(new[] { "Doe", "[F]linders", "Roe" }.Select(value => html ? "<div class=\"csl-entry\">" + value + "</div>" : value), result.Bibliography.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void QuotesAndCommasDoNotChangeTitleOrder(bool macro) {
        const string json = """
            [{"id":"c","type":"book","title":"Simple 'title' here C"},
             {"id":"a","type":"book","title":"Simple title, here A"},
             {"id":"b","type":"article-journal","title":"Simple title here B"}]
            """;
        Assert.Equal(new[] { "a", "b", "c" }, Render(json, macro ? "macro=\"title\"" : "variable=\"title\"", "<text variable=\"title\"/>").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PunctuationOnlyChunksBeforeAndBetweenNumbersDoNotShiftNumericComparisons(bool descending) {
        string json = JsonSerializer.Serialize(new[] {
            new { id = "ten", type = "book", title = "[10] Z" },
            new { id = "two", type = "book", title = "2 A" },
            new { id = "eleven", type = "book", title = "(11) B" }
        });
        Assert.Equal(descending ? new[] { "eleven", "ten", "two" } : new[] { "two", "ten", "eleven" },
            Render(json, "variable=\"title\"", "<text variable=\"title\"/>", descending: descending).Bibliography.Select(entry => entry.Key));
    }

    [Fact]
    public void NumericRunsRemainSeparateWhenPunctuationSeparatesThem() {
        const string json = """
            [{"id":"twelve","type":"book","title":"Volume 12"},
             {"id":"one-two","type":"book","title":"Volume 1.2"},
             {"id":"two","type":"book","title":"Volume 2"}]
            """;
        Assert.Equal(new[] { "one-two", "two", "twelve" }, Render(json, "variable=\"title\"", "<text variable=\"title\"/>").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("fr-FR")]
    public void PunctuationAndAccentTiesUseTheNextKeyAndKeepStableOrder(string locale) {
        const string json = """
            [{"id":"late","type":"book","title":"Éclair!","volume":"2"},
             {"id":"early","type":"book","title":"eclair","volume":"1"},
             {"id":"same","type":"book","title":"[Eclair]","volume":"1"}]
            """;
        Assert.Equal(new[] { "early", "same", "late" }, Render(json, "variable=\"title\"", "<text variable=\"title\"/>", extraKeys: "<key variable=\"volume\"/>", locale: locale).Bibliography.Select(entry => entry.Key));
    }

    [Fact]
    public void ContributorBoundariesStayDistinctFromIgnoredPunctuation() {
        const string json = """
            [{"id":"joined","type":"book","author":[{"family":"Doe Zeta"}]},
             {"id":"split","type":"book","author":[{"family":"Doe"},{"family":"Zeta"}]}]
            """;
        Assert.Equal(new[] { "split", "joined" }, Render(json, "variable=\"author\"", "<names variable=\"author\"><name/></names>").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("en-US", false)]
    [InlineData("en-US", true)]
    [InlineData("da-DK", false)]
    [InlineData("da-DK", true)]
    public void InternalWordSpacesRemainSignificantWhenPunctuationIsIgnored(string locale, bool macro) {
        const string json = """
            [{"id":"ab","type":"book","title":"Ab Delrahman"},
             {"id":"hansen","type":"book","title":"A Hansen"},
             {"id":"er","type":"book","title":"[A er]"}]
            """;
        Assert.Equal(new[] { "er", "hansen", "ab" }, Render(json, macro ? "macro=\"title\"" : "variable=\"title\"", "<text variable=\"title\"/>", locale: locale).Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("x")]
    [InlineData("😀")]
    public void LongSortRunsConsumeTheWorkBudgetForBmpAndSupplementaryText(string unit) {
        string tail = string.Concat(Enumerable.Repeat(unit, 32768));
        string json = JsonSerializer.Serialize(new[] { new { id = "a", type = "book", title = "A" + tail }, new { id = "b", type = "book", title = "B" + tail } });
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text value=\"Cite\"/></layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><text value=\"Entry\"/></layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style, new CslRenderOptions { MaximumRenderingOperations = 100 });
        Assert.Throws<InvalidDataException>(() => processor.Render(Array.Empty<CslCitation>(), includeUncitedItems: true));
    }

    private static CslRenderResult Render(string json, string key, string layout, bool html = false, bool descending = false, string extraKeys = "", string locale = "en-US") {
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\" default-locale=\"" + locale + "\"><macro name=\"names\"><names variable=\"author\"><name/></names></macro><macro name=\"title\"><text variable=\"title\" quotes=\"true\"/></macro><citation><layout><text variable=\"title\"/></layout></citation><bibliography><sort><key " + key + " sort=\"" + (descending ? "descending" : "ascending") + "\"/>" + extraKeys + "</sort><layout>" + layout + "</layout></bibliography></style>");
        return new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(Array.Empty<CslCitation>(), includeUncitedItems: true);
    }
}
