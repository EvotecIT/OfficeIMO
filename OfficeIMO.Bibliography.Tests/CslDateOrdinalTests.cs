namespace OfficeIMO.Bibliography.Tests;

public sealed class CslDateOrdinalTests {
    [Theory]
    [InlineData("[2024,1,1]", "1m")]
    [InlineData("[2024,2,1]", "1f")]
    [InlineData("[2024,3,1]", "1n")]
    [InlineData("[2024]", "")]
    public void OrdinalDaysUseTheMonthGenderEvenWhenTheMonthIsNotPrinted(string parts, string expected) =>
        Assert.Equal(expected, Render("[" + parts + "]"));

    [Fact]
    public void EachEndpointOfADateRangeUsesItsOwnMonthGender() =>
        Assert.Equal("1m January–2f February", Render("[[2024,1,1],[2024,2,2]]", includeMonth: true));

    [Theory]
    [InlineData("[[2024,1,1]]", "1m")]
    [InlineData("[[2024,1,2]]", "2")]
    [InlineData("[[2024,1,1],[2024,1,14]]", "1m–14")]
    [InlineData("[[2024,1,14],[2024,2,1]]", "14–1f")]
    public void TheDayOneOptionAppliesIndependentlyToTypedDateEndpoints(string parts, string expected) =>
        Assert.Equal(expected, Render(parts, limitDayOne: true));

    [Fact]
    public void APartialModernOrdinalSetAlsoFormatsDateDays() =>
        Assert.Equal("11x", Render("[[2024,1,11]]", ordinalTerms: "<term name=\"ordinal-11\">x</term>"));

    private static string Render(string parts, bool includeMonth = false, bool limitDayOne = false, string? ordinalTerms = null) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"issued\":{\"date-parts\":" + parts + "}}]", BibliographyFormat.CslJson).Document;
        string month = includeMonth ? "<date-part name=\"month\"/>" : string.Empty;
        CslStyle style = CslStyle.Parse($"""
            <style xmlns="http://purl.org/net/xbiblio/csl" version="1.0" class="in-text">
              <locale><style-options limit-day-ordinals-to-day-1="{limitDayOne.ToString().ToLowerInvariant()}"/><terms>
                <term name="month-01" gender="masculine">January</term>
                <term name="month-02" gender="feminine">February</term>
                <term name="month-03">March</term>
                {ordinalTerms ?? "<term name=\"ordinal\">n</term><term name=\"ordinal\" gender-form=\"masculine\">m</term><term name=\"ordinal\" gender-form=\"feminine\">f</term>"}
              </terms></locale>
              <citation><layout><date variable="issued" delimiter=" "><date-part name="day" form="ordinal"/>{month}</date></layout></citation>
            </style>
            """);
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style).Render(new[] { citation }).Citations.Single().Content;
    }
}
