using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordAllSeverityBatch19SecurityTests {
    [Fact]
    public void AutomaticColorSentinelRemainsLowercaseAcrossPublicColorSetters() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("automatic");
        paragraph.RunColorValue = "AUTO";
        document.Borders.LeftStyle = WordBorderStyle.Single;
        document.Borders.LeftColorValue = "Auto";
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Borders.TopStyle = WordBorderStyle.Single;
        cell.Borders.TopColorValue = "aUtO";

        Assert.Equal("auto", paragraph.RunColorValue);
        Assert.Null(paragraph.Color);
        Assert.Equal("auto", document.Borders.LeftColorValue);
        Assert.Equal("auto", cell.Borders.TopColorValue);
        Assert.Empty(new OpenXmlValidator().Validate(document._wordprocessingDocument));
    }
}
