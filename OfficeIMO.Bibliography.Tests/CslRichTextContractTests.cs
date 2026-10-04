using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslRichTextContractTests {
    [Theory]
    [InlineData("font-variant: small-caps;")]
    [InlineData(" FONT-VARIANT : SMALL-CAPS ")]
    [InlineData("color:red; font-variant:\tsmall-caps; background:url(https://example.org)")]
    public void SmallCapsInputAcceptsCssWhitespaceWithoutForwardingOtherDeclarations(string css) {
        string input = "<span style=\"" + css + "\">Here</span>";
        Assert.Equal("<span style=\"font-variant:small-caps;\">Here</span>", Render(input, CslOutputFormat.Html));
        Assert.Equal("Here", Render(input, CslOutputFormat.PlainText));
    }

    [Fact]
    public void SmallCapsInputRetainsItsCaseAndFlipsEnclosingSmallCaps() {
        Assert.Equal("<span style=\"font-variant:small-caps;\">Before <span style=\"font-variant:normal;\">iPhone</span> After</span>",
            Render("Before <span style=\"font-variant: small-caps;\">iPhone</span> After", CslOutputFormat.Html,
                "font-variant=\"small-caps\" text-case=\"capitalize-all\""));
    }

    [Fact]
    public void SmallCapsAndCaseProtectionOnTheSameSpanRetainBothSemantics() {
        Assert.Equal("<span style=\"font-variant:small-caps;\">iPhone</span>",
            Render("<span class=\"nocase\" style=\"font-variant: small-caps;\">iPhone</span>",
                CslOutputFormat.Html, "text-case=\"capitalize-all\""));
    }

    [Theory]
    [InlineData("nocase\tnodecor")]
    [InlineData("nodecor nocase")]
    public void CombinedProtectionClassesRetainOwnFormattingAndResetEnclosingEmphasis(string classes) {
        Assert.Equal("<i>Before <span style=\"font-variant:small-caps;font-style:normal;\">iPhone</span> After</i>",
            Render("Before <span class=\"" + classes + "\" style=\"font-variant: small-caps;\">iPhone</span> After",
                CslOutputFormat.Html, "font-style=\"italic\" text-case=\"capitalize-all\""));
    }

    [Theory]
    [InlineData("nodecor")]
    [InlineData("nocase nodecor")]
    [InlineData("nodecor\tnocase")]
    public void DecorationResetPrecedesOwnSmallCapsWhenTheEnclosingStyleUsesSmallCaps(string classes) {
        Assert.Equal("<span style=\"font-variant:small-caps;\"><span style=\"font-variant:small-caps;\">iPhone</span></span>",
            Render("<span class=\"" + classes + "\" style=\"font-variant: small-caps;\">iPhone</span>",
                CslOutputFormat.Html, "font-variant=\"small-caps\" text-case=\"capitalize-all\""));
    }

    [Theory]
    [InlineData("A &amp; B <i>C &lt; D &copy;</i>", "A & B C < D ©", "A &amp; B <i>C &lt; D &#169;</i>")]
    [InlineData("&lt;i&gt;literal&lt;/i&gt; <b>bold</b>", "<i>literal</i> bold", "&lt;i&gt;literal&lt;/i&gt; <b>bold</b>")]
    [InlineData("&#60;i&#62;literal&#60;/i&#62; <b>bold</b>", "<i>literal</i> bold", "&lt;i&gt;literal&lt;/i&gt; <b>bold</b>")]
    [InlineData("A &unknown; <i>B</i>", "A &unknown; B", "A &amp;unknown; <i>B</i>")]
    [InlineData("<i title=\"A &quot;quoted&quot; attribute\">safe</i>", "safe", "<i>safe</i>")]
    [InlineData("&lt;script&gt;bad()&lt;/script&gt; <i>safe</i>", "<script>bad()</script> safe", "&lt;script&gt;bad()&lt;/script&gt; <i>safe</i>")]
    public void EntitiesRemainTextWhileAdjacentInputMarkupKeepsItsFormatting(string input, string plain, string html) {
        Assert.Equal(plain, Render(input, CslOutputFormat.PlainText));
        Assert.Equal(html, Render(input, CslOutputFormat.Html));
    }

    private static string Render(string title, CslOutputFormat format, string formatting = "") {
        string json = JsonSerializer.Serialize(new[] { new { id = "one", type = "book", title } });
        var document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"title\" " + formatting + "/></layout></citation></style>");
        var citation = new CslCitation("cluster");
        citation.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style, new CslRenderOptions { OutputFormat = format }).Render(new[] { citation }).Citations.Single().Content;
    }
}
