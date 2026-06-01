using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Linq;
using System.Text;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Simple_Shapes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeShapes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeShapes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native shapes");
                document.AddShape(ShapeType.Rectangle, 100, 36, "#ccb399", "#1a334d", 2.5);
                document.AddShape(ShapeType.Ellipse, 60, 30, "#d0e6ff", "#224466", 1.25);
                document.AddShape(ShapeType.Line, 80, 0, "#ffffff", "#008000", 2);
                document.AddParagraph("After native shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before native shapes", allText);
                Assert.Contains("After native shapes", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("0.8 0.702 0.6 rg", content);
            Assert.Contains("0.102 0.2 0.302 RG", content);
            Assert.Contains("2.5 w", content);
            Assert.Contains(" re B", content);
            Assert.Contains("0 0.502 0 RG", content);
            Assert.Contains("2 w", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_DrawingML_Preset_Shapes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingPresetShapes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingPresetShapes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native DrawingML shapes");
                WordShape triangle = document.AddParagraph().AddShapeDrawing(ShapeType.Triangle, 72, 48);
                triangle.FillColorHex = "#99ccff";
                triangle.StrokeColorHex = "#003366";
                triangle.StrokeWeight = 1.5;
                Assert.Equal(1.5, triangle.StrokeWeight);
                WordShape diamond = document.AddParagraph().AddShapeDrawing(ShapeType.Diamond, 64, 48);
                diamond.FillColorHex = "#ffe699";
                WordShape arrow = document.AddParagraph().AddShapeDrawing(ShapeType.RightArrow, 96, 36);
                arrow.FillColorHex = "#b7e1cd";
                WordShape star = document.AddParagraph().AddShapeDrawing(ShapeType.Star5, 56, 56);
                star.FillColorHex = "#f4b183";
                document.AddParagraph("After native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before native DrawingML shapes", allText);
                Assert.Contains("After native DrawingML shapes", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("0.6 0.8 1 rg", content);
            Assert.Contains("0 0.2 0.4 RG", content);
            Assert.Contains("1.5 w", content);
            Assert.Contains("1 0.902 0.6 rg", content);
            Assert.Contains("0.718 0.882 0.804 rg", content);
            Assert.Contains("0.957 0.694 0.514 rg", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Remaining_DrawingML_Preset_Shapes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRemainingDrawingPresetShapes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRemainingDrawingPresetShapes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before remaining native DrawingML shapes");
                document.AddParagraph().AddShapeDrawing(ShapeType.Heart, 64, 56).FillColorHex = "#ff0000";
                document.AddParagraph().AddShapeDrawing(ShapeType.Cloud, 88, 52).FillColorHex = "#00ff00";
                document.AddParagraph().AddShapeDrawing(ShapeType.Donut, 64, 64).FillColorHex = "#0000ff";
                document.AddParagraph().AddShapeDrawing(ShapeType.Can, 64, 64).FillColorHex = "#ffff00";
                document.AddParagraph().AddShapeDrawing(ShapeType.Cube, 72, 60).FillColorHex = "#00ffff";
                document.AddParagraph("After remaining native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before remaining native DrawingML shapes", allText);
                Assert.Contains("After remaining native DrawingML shapes", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("1 0 0 rg", content);
            Assert.Contains("0 1 0 rg", content);
            Assert.Contains("0 0 1 rg", content);
            Assert.Contains("1 1 0 rg", content);
            Assert.Contains("0 1 1 rg", content);
            Assert.Contains(" h\n", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Simple_TextBoxes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTextBoxes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTextBoxes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native text box");
                WordTextBox textBox = document.AddTextBox("Native text box body");
                textBox.WidthCentimeters = 7;
                textBox.HorizontalAlignment = WordHorizontalAlignmentValues.Center;
                textBox.Paragraphs[0].Bold = true;
                document.AddParagraph("After native text box");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before native text box", allText);
                Assert.Contains("Native text box body", allText);
                Assert.Contains("After native text box", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains(" re S", content);
        }
    }
}
