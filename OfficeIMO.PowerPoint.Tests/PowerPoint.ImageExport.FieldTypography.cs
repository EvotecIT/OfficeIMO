using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ImageExportPreservesFieldOnlyTypographyForTextBoxesAndTableCells() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(340, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointParagraph textBoxParagraph = slide.AddTextBoxPoints("seed", 20, 20, 135, 60).Paragraphs[0];
            textBoxParagraph.Paragraph.RemoveAllChildren<A.Run>();
            textBoxParagraph.AddField("text field", "custom", null, run => {
                run.Bold = true;
                run.Italic = true;
                run.FontName = "Arial";
                run.FontSize = 15;
                run.UnderlineStyle = PowerPointUnderlineStyle.Double;
                run.Capitalization = PowerPointCapitalization.AllCaps;
            });

            PowerPointTableCell cell = slide.AddTablePoints(1, 1, 175, 20, 140, 55).GetCell(0, 0);
            cell.Text = "seed";
            PowerPointParagraph cellParagraph = cell.Paragraphs[0];
            cellParagraph.Paragraph.RemoveAllChildren<A.Run>();
            cellParagraph.AddField("table field", "custom", null, run => {
                run.Bold = true;
                run.FontName = "Consolas";
                run.FontSize = 13;
                run.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
                run.Capitalization = PowerPointCapitalization.SmallCaps;
            });

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            OfficeDrawingRichText textBoxText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>(),
                element => element.PlainText.Contains("TEXT FIELD", StringComparison.Ordinal));
            OfficeRichTextRun textBoxRun = Assert.Single(textBoxText.Runs);
            Assert.True(textBoxRun.Bold);
            Assert.True(textBoxRun.Italic);
            Assert.Equal("Arial", textBoxRun.FontFamily);
            Assert.Equal(OfficeTextDecorationStyle.Double, textBoxRun.UnderlineStyle);

            OfficeDrawingRichText tableText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>(),
                element => element.PlainText.Contains("TABLE FIELD", StringComparison.Ordinal));
            OfficeRichTextRun tableRun = Assert.Single(tableText.Runs);
            Assert.True(tableRun.Bold);
            Assert.Equal("Consolas", tableRun.FontFamily);
            Assert.Equal(OfficeTextDecorationStyle.Wavy, tableRun.UnderlineStyle);

            Assert.Contains("TEXT FIELD", svgText, StringComparison.Ordinal);
            Assert.Contains("TABLE FIELD", svgText, StringComparison.Ordinal);
            Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
        }
    }
}
