using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlWordTablePadding {
    [Theory]
    [InlineData("padding:9px 8px", "135", "120", "135", "120")]
    [InlineData("padding:1pt 2pt 3pt 4pt;padding-left:0", "20", "40", "60", "0")]
    [InlineData("padding:40000pt", "32767", "32767", "32767", "32767")]
    [InlineData("font-size:24px;padding:1em", "360", "360", "360", "360")]
    [InlineData("padding:1rem", "300", "300", "300", "300")]
    [InlineData("direction:rtl;padding:1px;padding-inline-start:8px", "15", "120", "15", "15")]
    public void CssCellPaddingBecomesBoundedNativeMargins(string css, string top, string right, string bottom, string left) {
        HtmlToWordResult result = HtmlConversionDocument.Parse(
            $"<style>html {{font-size:20px}} th,td {{{css}}}</style><table><tr><th>Header</th><td>Value</td></tr></table>")
            .ToWordDocumentResult();
        using WordDocument word = result.RequireValue();
        using MemoryStream artifact = word.ToStream();
        using WordprocessingDocument package = WordprocessingDocument.Open(artifact, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToArray();
        Assert.Equal(2, cells.Length);
        foreach (TableCell cell in cells) {
            TableCellMargin margins = Assert.IsType<TableCellMargin>(cell.TableCellProperties?.TableCellMargin);
            Assert.Equal(top, margins.TopMargin?.Width?.Value);
            Assert.Equal(right, margins.RightMargin?.Width?.Value);
            Assert.Equal(bottom, margins.BottomMargin?.Width?.Value);
            Assert.Equal(left, margins.LeftMargin?.Width?.Value);
        }
        if (css.Contains("40000", StringComparison.Ordinal)) {
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("native margin limit", StringComparison.Ordinal));
        } else {
            Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "UnsupportedCssDeclaration"
                && (diagnostic.Source ?? string.Empty).Contains("padding", StringComparison.Ordinal));
        }
    }
}
