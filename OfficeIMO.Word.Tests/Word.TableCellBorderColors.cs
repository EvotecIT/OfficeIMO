using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class WordTableCellBorderColors {
    [Fact]
    public void ClearingCellBorderColorsRemovesAttributesAndRetainsBorders() {
        using var document = WordDocument.Create();
        var borders = document.AddTable(1, 1).Rows[0].Cells[0].Borders;
        borders.LeftStyle = borders.RightStyle = borders.TopStyle = borders.BottomStyle =
            borders.InsideHorizontalStyle = borders.InsideVerticalStyle = borders.StartStyle = borders.EndStyle =
            borders.TopLeftToBottomRightStyle = borders.TopRightToBottomLeftStyle = WordBorderStyle.Single;
        borders.LeftColorHex = borders.RightColorHex = borders.TopColorHex = borders.BottomColorHex =
            borders.InsideHorizontalColorHex = borders.InsideVerticalColorHex = borders.StartColorHex = borders.EndColorHex =
            borders.TopLeftToBottomRightColorHex = borders.TopRightToBottomLeftColorHex = "#123456";
        borders.LeftColorHex = borders.RightColorHex = borders.TopColorHex = borders.BottomColorHex =
            borders.InsideHorizontalColorHex = borders.InsideVerticalColorHex = borders.StartColorHex = borders.EndColorHex =
            borders.TopLeftToBottomRightColorHex = borders.TopRightToBottomLeftColorHex = null;

        using var artifact = document.ToStream();
        using var package = WordprocessingDocument.Open(artifact, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package));
        var savedBorders = Assert.Single(package.MainDocumentPart!.Document.Descendants<TableCellBorders>());
        Assert.Equal(10, savedBorders.ChildElements.Count);
        foreach (var border in savedBorders.ChildElements.Cast<BorderType>()) {
            Assert.Null(border.Color);
            Assert.Equal(BorderValues.Single, border.Val!.Value);
        }
    }
}
