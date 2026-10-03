namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyRichTextConversionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void CrossFormatExportReportsLiteralCslRichTextSemantics(BibliographyFormat target) {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"title\":\"<i>Italic</i> &amp; meaning\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        Assert.Equal(source, document.Write().Content);
        BibliographyWriteResult result = document.Write(new BibliographyWriteOptions { Format = target, Mode = BibliographyWriterMode.Canonical });
        Assert.True(result.Report.HasLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV252" && diagnostic.Field == "title");
        Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = target, RequireNoLoss = true }));
        Assert.False(document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }).Report.HasLoss);
    }

    [Fact]
    public void BibCaseProtectionIsExplicitWhenExportedToCsl() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={A {DNA} study}}", BibliographyFormat.BibLatex).Document;
        BibliographyWriteResult result = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson });
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV252" && diagnostic.Field == "title");
        Assert.Equal("A {DNA} study", document.Items[0].Title);
    }
}
