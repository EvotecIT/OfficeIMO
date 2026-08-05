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
        paragraph.ColorHex = "AUTO";
        document.Borders.LeftStyle = WordBorderStyle.Single;
        document.Borders.LeftColorHex = "Auto";
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Borders.TopStyle = WordBorderStyle.Single;
        cell.Borders.TopColorHex = "aUtO";

        Assert.Equal("auto", paragraph.ColorHex);
        Assert.Equal("auto", document.Borders.LeftColorHex);
        Assert.Equal("auto", cell.Borders.TopColorHex);
        Assert.Empty(new OpenXmlValidator().Validate(document._wordprocessingDocument));
    }
}
