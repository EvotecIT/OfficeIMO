using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityBatch16SecurityTests {
    [Fact]
    public void WordToHtmlDropsInvalidTableColorTokens() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.ShadingFillColorHex = "ffffff;background-image:url(https://tracker.invalid/pixel)";
        cell.Borders.LeftStyle = OfficeIMO.Word.WordBorderStyle.Single;
        cell.Borders.LeftColorHex = "ffffff;position:fixed";

        string html = document.ToHtml();

        Assert.DoesNotContain("background-image", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("position:fixed", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("border-left:1px solid black", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordToHtmlQuotesDocumentControlledRunFonts() {
        const string hostileFont = "Arial;background-image:url(https://tracker.invalid/pixel)";
        using WordDocument document = WordDocument.Create();
        WordParagraph plain = document.AddParagraph("plain");
        plain.FontFamily = hostileFont;
        WordParagraph equationAdjacent = document.AddParagraph("before ");
        equationAdjacent.FontFamily = hostileFont;
        equationAdjacent.AddEquation("<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>x</m:t></m:r></m:oMath>");

        string html = document.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });

        Assert.DoesNotContain(";background-image:url", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("font-family:", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordToHtmlIgnoresMissingAndDuplicateStyleIds() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style { Type = StyleValues.Paragraph, StyleId = "Collision" },
            new Style { Type = StyleValues.Paragraph, StyleId = "collision" },
            new Style { Type = StyleValues.Paragraph });
        document.AddParagraph("styled").SetStyleId("Collision");

        Exception? exception = Record.Exception(() =>
            document.ToHtml(new WordToHtmlOptions { IncludeParagraphClasses = true }));

        Assert.Null(exception);
    }

    [Fact]
    public void WordToHtmlKeepsDistinctParagraphsWithEqualPublicValues() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("same-content");
        document.AddParagraph("same-content");

        string html = document.ToHtml();

        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(html, ">same-content</p>").Count);
    }

    [Fact]
    public void WordToHtmlExpandsDeepTableCellWrappersWithoutRecursion() {
        const int depth = 5_000;
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell._tableCell.RemoveAllChildren();
        OpenXmlCompositeElement parent = cell._tableCell;
        for (int index = 0; index < depth; index++) {
            var content = new SdtContentBlock();
            parent.Append(new SdtBlock(new SdtProperties(), content));
            parent = content;
        }
        parent.Append(new Paragraph(new Run(new Text("deep-table-cell"))));

        string html = document.ToHtml(new WordToHtmlOptions {
            MaxDocumentElements = 100_000
        });

        Assert.Contains("deep-table-cell", html, StringComparison.Ordinal);
    }

    [Fact]
    public void WordToHtmlRejectsDeepOmmlBeforeRecursiveProjection() {
        const int depth = 5_000;
        using WordDocument document = WordDocument.Create();
        var equation = new M.OfficeMath();
        OpenXmlCompositeElement parent = equation;
        for (int index = 0; index < depth; index++) {
            var numerator = new M.Numerator();
            parent.Append(new M.Fraction(
                numerator,
                new M.Denominator(new M.Run(new M.Text("1")))));
            parent = numerator;
        }
        parent.Append(new M.Run(new M.Text("x")));
        document.AddParagraph()._paragraph.Append(equation);

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
            document.ToHtmlResult(new WordToHtmlOptions {
                MaxDocumentElements = 100_000,
                MaxEquationNestingDepth = 64
            }));

        Assert.Equal("WordEquationDepthLimitExceeded", exception.Code);
        Assert.Equal("EquationOmml", exception.LimitSource);
    }
}
