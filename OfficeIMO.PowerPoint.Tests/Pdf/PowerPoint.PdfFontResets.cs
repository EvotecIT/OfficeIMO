using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfFontResetTests {
    [Theory]
    [InlineData("run")]
    [InlineData("textbox")]
    [InlineData("cell")]
    public void Explicit_regular_font_overrides_inherited_bold_and_italic_after_save(string surface) {
        using var saved = new MemoryStream();
        using (var presentation = PowerPointPresentation.Create()) {
            presentation.SlideSize.SetSizePoints(260, 180);
            var slide = presentation.AddSlide();
            PowerPointParagraph paragraph;
            if (surface == "cell") {
                var table = slide.AddTablePoints(1, 1, 30, 34, 160, 80);
                table.ApplyStyle(PowerPointTableStylePreset.Plain);
                var cell = table.GetCell(0, 0);
                cell.Text = "Plain"; cell.FontName = "Helvetica"; cell.FontSize = 18;
                paragraph = cell.Paragraphs[0];
                paragraph.Paragraph.ParagraphProperties = new A.ParagraphProperties(new A.DefaultRunProperties { Bold = true, Italic = true });
                cell.Bold = false; cell.Italic = false;
            } else {
                var box = slide.AddTextBoxPoints("Plain", 30, 34, 160, 80);
                box.FontName = "Helvetica"; box.FontSize = 18;
                paragraph = box.Paragraphs[0];
                paragraph.Paragraph.ParagraphProperties = new A.ParagraphProperties(new A.DefaultRunProperties { Bold = true, Italic = true });
                if (surface == "run") { paragraph.Runs[0].Bold = false; paragraph.Runs[0].Italic = false; }
                else { box.Bold = false; box.Italic = false; }
            }
            presentation.Save(saved);
        }
        saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        var actual = surface == "cell" ? reopened.Slides[0].Tables.Single().GetCell(0, 0).Paragraphs[0]
            : reopened.Slides[0].TextBoxes.Single().Paragraphs[0];
        var properties = Assert.Single(actual.Paragraph.Elements<A.Run>()).RunProperties!;
        Assert.NotNull(properties.Bold); Assert.False(properties.Bold.Value);
        Assert.NotNull(properties.Italic); Assert.False(properties.Italic.Value);
        using var pdf = PdfPigDocument.Open(reopened.ToPdfBytes());
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("Plain", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            Assert.DoesNotContain("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oblique", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Italic", letter.FontName, StringComparison.OrdinalIgnoreCase);
        });
    }
}
