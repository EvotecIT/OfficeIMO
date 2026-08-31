namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave37RegressionTests {
    [Theory]
    [InlineData("% first\r% second\r@book{x,title={Title}}", BibliographyFormat.BibLatex)]
    [InlineData("// first\r// second\r[{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    [InlineData("% first\r\n@book{x,title={Title}}", BibliographyFormat.BibLatex)]
    [InlineData("// first\r\n[{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    public void Detection_accepts_line_comments_with_CR_or_CRLF_endings(string source, BibliographyFormat expected) {
        Assert.Equal(expected, BibliographyDocument.Parse(source).Document.SourceFormat);
    }

    [Fact]
    public void Canonical_CR_only_Bib_output_with_a_leading_comment_is_auto_detectable() {
        BibliographyDocument document = BibliographyDocument.Parse("% retained\n@book{x,title={Before}}", BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, LineEnding = "\r" });
        BibliographyReadResult reopened = BibliographyDocument.Parse(written.Content);

        Assert.Equal(BibliographyFormat.BibLatex, reopened.Document.SourceFormat);
        Assert.Equal("After", Assert.Single(reopened.Document.Items).Title);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void CSL_syntax_diagnostics_map_all_supported_line_endings(string lineEnding) {
        string source = "[" + lineEnding + "{\"title\":\"Ł\",?}]";
        int expectedOffset = source.IndexOf('?');
        int lineStart = lineEnding.Length + 1;

        BibliographyDiagnostic diagnostic = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Diagnostics, value => value.Code == "BIBCSL002");

        Assert.Equal(expectedOffset, diagnostic.Offset);
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(expectedOffset - lineStart + 1, diagnostic.Column);
    }

    [Fact]
    public void Undefined_writer_modes_are_rejected_before_preserve_or_normalization() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Exact}}", BibliographyFormat.BibLatex).Document;
        var options = new BibliographyWriteOptions { Mode = (BibliographyWriterMode)99 };

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => document.Write(options));

        Assert.Equal(nameof(BibliographyWriteOptions.Mode), exception.ParamName);
    }

    [Fact]
    public void Undefined_destination_formats_are_rejected_with_the_other_writer_enums() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var options = new BibliographyWriteOptions { Format = (BibliographyFormat)99 };

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => document.Write(options));

        Assert.Equal(nameof(BibliographyWriteOptions.Format), exception.ParamName);
    }

    [Theory]
    [InlineData("[2024,0]")]
    [InlineData("[2024,13]")]
    [InlineData("[2024,1,0]")]
    [InlineData("[2024,1,32]")]
    [InlineData("[2024,1,1],[2025,13,1]")]
    [InlineData("[2024,1,1],[2025,1,32]")]
    public void Out_of_range_CSL_date_parts_remain_native_and_reopen_exactly(string dateParts) {
        string source = "[{\"id\":\"x\",\"type\":\"book\",\"issued\":{\"date-parts\":[" + dateParts + "]}}]";
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson);
        BibliographyDate date = Assert.Single(read.Document.Items[0].Dates);
        read.Document.Items[0].Title = "Edited";

        BibliographyWriteResult written = read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Dates);

        Assert.Null(date.Month);
        Assert.Null(date.Day);
        Assert.Contains(date.NativeFields, field => field.Name == "date-parts");
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL005");
        Assert.Contains(reopened.NativeFields, field => field.Name == "date-parts" && RemoveWhitespace(field.Value).Contains(dateParts, StringComparison.Ordinal));
    }

    private static string RemoveWhitespace(string value) => new string(value.Where(static character => !char.IsWhiteSpace(character)).ToArray());
}
