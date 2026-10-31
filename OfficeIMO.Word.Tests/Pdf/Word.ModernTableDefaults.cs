using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", "header")]
    [InlineData("body", "body")]
    [InlineData("body", "footer")]
    [InlineData("nested", "header")]
    [InlineData("nested", "body")]
    [InlineData("nested", "footer")]
    [InlineData("header", "header")]
    [InlineData("header", "body")]
    [InlineData("header", "footer")]
    [InlineData("footer", "header")]
    [InlineData("footer", "body")]
    [InlineData("footer", "footer")]
    public void SaveAsPdf_ModernTablesRetainDirectCellTextOverrides(string scope, string role) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body marker");
        document.AddHeadersAndFooters();
        WordTable table = scope switch {
            "body" => document.AddTable(2, 2),
            "nested" => document.AddTable(1, 1).Rows[0].Cells[0].AddTable(2, 2),
            "header" => document.Header.Default!.AddTable(2, 2),
            "footer" => document.Footer.Default!.AddTable(2, 2),
            _ => throw new ArgumentOutOfRangeException(nameof(scope))
        };
        int row = role == "header" ? 0 : 1;
        if (role == "footer") {
            table.ConditionalFormattingLastRow = true;
            GetStyleDefinition(document, "OfficeIMOModern").AppendChild(new W.TableStyleProperties(
                new W.RunPropertiesBaseStyle(new W.Bold { Val = true })) { Type = W.TableStyleOverrideValues.LastRow });
        }
        WordTableCell cell = table.Rows[row].Cells[0];
        GetStyleDefinition(document, "OfficeIMOModern").StyleRunProperties =
            new W.StyleRunProperties(new W.Italic { Val = true });
        WordParagraph text = cell.Paragraphs[0];
        text.Text = "OVERRIDE";
        text.FontSize = OfficeLength.Points(18);
        text.ColorHex = "008000";
        text.Bold = false;
        text.Italic = false;
        text.ParagraphAlignment = WordParagraphAlignment.Right;
        cell.ShadingFillColorHex = "FF0000";
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
        document.CleanupDocument();
        using var saved = WordDocument.Load(new MemoryStream(document.ToBytes()));
        W.RunProperties authored = saved._wordprocessingDocument.MainDocumentPart!.Document
            .Descendants<W.Run>().SingleOrDefault(run => run.InnerText == "OVERRIDE")?.RunProperties!;
        if (scope is "header" or "footer") {
            var part = saved._wordprocessingDocument.MainDocumentPart!;
            authored = (scope == "header" ? part.HeaderParts.SelectMany(item => item.Header.Descendants<W.Run>())
                : part.FooterParts.SelectMany(item => item.Footer.Descendants<W.Run>()))
                .Single(run => run.InnerText == "OVERRIDE").RunProperties!;
        }
        Assert.False(authored.Bold!.Val!.Value);
        Assert.False(authored.BoldComplexScript!.Val!.Value);
        Assert.False(authored.Italic!.Val!.Value);
        Assert.False(authored.ItalicComplexScript!.Val!.Value);
        using var model = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeLength.Points(100)) });
        using var pdf = PdfPigDocument.Open(model.ToBytes());
        var page = pdf.GetPage(1);
        var word = Assert.Single(page.GetWords(), word => word.Text == "OVERRIDE");
        Assert.True(word.BoundingBox.Left > 150D);
        Assert.All(word.Letters, letter => {
            Assert.Equal(18D, letter.FontSize, 2);
            var rgb = letter.Color.ToRGBValues();
            Assert.Equal(0D, rgb.r, 2);
            Assert.Equal(128D / 255D, rgb.g, 2);
            Assert.Equal(0D, rgb.b, 2);
            Assert.DoesNotContain("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Italic", letter.FontName, StringComparison.OrdinalIgnoreCase);
        });
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() is var rgb &&
            rgb.r > .99D && rgb.g < .01D && rgb.b < .01D);
    }
}
