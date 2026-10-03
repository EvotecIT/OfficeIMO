namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyProtectedNameTests {
    [Theory]
    [InlineData("John {van der Waals}", "John", "van der Waals")]
    [InlineData("{de la Cruz}, Maria", "Maria", "de la Cruz")]
    [InlineData("{de la Cruz}, {Maria Elena}", "Maria Elena", "de la Cruz")]
    public void ProtectedNameComponentsRemainWholeThroughCslAndCanonicalBib(string input, string given, string family) {
        BibliographyDocument source = BibliographyDocument.Parse("@book{x, author={" + input + "}}", BibliographyFormat.BibLatex).Document;
        BibliographyName name = Assert.Single(Assert.Single(source.Items).Contributors).Name;
        Assert.Equal(given, name.Given);
        Assert.Equal(family, name.Family);
        Assert.Null(name.NonDroppingParticle);
        Assert.Null(name.Literal);
        foreach (BibliographyFormat format in new[] { BibliographyFormat.CslJson, BibliographyFormat.BibLatex }) {
            BibliographyWriteResult written = source.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
            BibliographyName reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items).Contributors).Name;
            Assert.Equal(given, reopened.Given);
            Assert.Equal(family, reopened.Family);
            Assert.Null(reopened.NonDroppingParticle);
        }
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData("\r\n")]
    public void CanonicalCslUsesRequestedLineEndingThroughout(string lineEnding) {
        BibliographyDocument source = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"title\":\"Text\",\"vendor\":{\"list\":[1,2]}}]", BibliographyFormat.CslJson).Document;
        BibliographyWriteResult written = source.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, LineEnding = lineEnding, RequireNoLoss = true });
        Assert.Contains(lineEnding, written.Content);
        Assert.DoesNotContain("\n", written.Content.Replace(lineEnding, string.Empty));
        Assert.DoesNotContain("\r", written.Content.Replace(lineEnding, string.Empty));
        Assert.False(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).HasErrors);
    }
}
