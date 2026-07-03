using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
        public void SaveAsPdf_OfficeIMOEngine_Preserves_DrawingML_Line_Preset_As_Horizontal() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingLinePresetGeometry.docx");
            using WordDocument document = WordDocument.Create(docPath);
            WordShape line = document.AddParagraph().AddShapeDrawing(ShapeType.Line, 80, 24);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeShape", BindingFlags.NonPublic | BindingFlags.Static)!;
            OfficeShape shape = Assert.IsType<OfficeShape>(method.Invoke(null, new object[] { line }));

            Assert.Equal(OfficeShapeKind.Line, shape.Kind);
            Assert.Equal(2, shape.Points.Count);
            Assert.Equal(shape.Points[0].Y, shape.Points[1].Y);
            Assert.Equal(0D, shape.Points[0].Y, precision: 3);
            Assert.True(shape.Points[1].X > shape.Points[0].X);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Suppresses_DrawingML_Line_NoFill_Stroke() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingLineNoFill.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingLineNoFill.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before no-fill line");
                document.AddParagraph().AddShapeDrawing(ShapeType.Line, 96, 1);
                document.AddParagraph("After no-fill line");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
                Wps.ShapeProperties shapeProperties = package.MainDocumentPart!.Document.Descendants<Wps.ShapeProperties>().First();
                A.Outline outline = shapeProperties.GetFirstChild<A.Outline>() ?? new A.Outline();
                if (outline.Parent == null) {
                    shapeProperties.Append(outline);
                }

                outline.RemoveAllChildren();
                outline.Append(new A.NoFill());
                package.MainDocumentPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(docPath)) {
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
            Assert.DoesNotContain(" RG", pageContent, System.StringComparison.Ordinal);
            Assert.DoesNotContain(" S", pageContent, System.StringComparison.Ordinal);
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
        public void SaveAsPdf_OfficeIMOEngine_Renders_Additional_DrawingML_Preset_Shapes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAdditionalDrawingPresetShapes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAdditionalDrawingPresetShapes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before additional native DrawingML shapes");
                document.AddParagraph().AddShapeDrawing(ShapeType.Parallelogram, 72, 40).FillColorHex = "#ff00ff";
                document.AddParagraph().AddShapeDrawing(ShapeType.Trapezoid, 72, 40).FillColorHex = "#ff9900";
                document.AddParagraph().AddShapeDrawing(ShapeType.Chevron, 80, 44).FillColorHex = "#808080";
                document.AddParagraph().AddShapeDrawing(ShapeType.Plus, 56, 56).FillColorHex = "#008080";
                document.AddParagraph().AddShapeDrawing(ShapeType.LeftRightArrow, 96, 40).FillColorHex = "#800000";
                document.AddParagraph("After additional native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before additional native DrawingML shapes", allText);
                Assert.Contains("After additional native DrawingML shapes", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("1 0 1 rg", content);
            Assert.Contains("1 0.6 0 rg", content);
            Assert.Contains("0.502 0.502 0.502 rg", content);
            Assert.Contains("0 0.502 0.502 rg", content);
            Assert.Contains("0.502 0 0 rg", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_DrawingML_Scheme_Fill_And_Dashed_Outline() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingSchemeShape.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingSchemeShape.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before scheme shape");
                document.AddParagraph().AddShapeDrawing(ShapeType.Rectangle, 96, 44);
                document.AddParagraph("After scheme shape");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
                Wps.ShapeProperties shapeProperties = package.MainDocumentPart!.Document.Descendants<Wps.ShapeProperties>().First();
                A.PresetGeometry? geometry = shapeProperties.GetFirstChild<A.PresetGeometry>();
                shapeProperties.RemoveAllChildren<A.SolidFill>();
                var fill = new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 });
                if (geometry != null) {
                    shapeProperties.InsertAfter(fill, geometry);
                } else {
                    shapeProperties.Append(fill);
                }

                A.Outline outline = shapeProperties.GetFirstChild<A.Outline>() ?? new A.Outline();
                if (outline.Parent == null) {
                    shapeProperties.Append(outline);
                }

                outline.Width = 25400;
                outline.RemoveAllChildren();
                outline.Append(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }));
                outline.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Dash });
                package.MainDocumentPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(docPath)) {
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("0.929 0.49 0.192 rg", content);
            Assert.Contains("0.267 0.447 0.769 RG", content);
            Assert.Contains("2 w", content);
            Assert.Contains("[6 3] 0 d", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_DrawingML_Gradient_Fill_And_Alpha() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingGradientShape.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingGradientShape.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before gradient shape");
                document.AddParagraph().AddShapeDrawing(ShapeType.Rectangle, 112, 48);
                document.AddParagraph("After gradient shape");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
                Wps.ShapeProperties shapeProperties = package.MainDocumentPart!.Document.Descendants<Wps.ShapeProperties>().First();
                A.PresetGeometry? geometry = shapeProperties.GetFirstChild<A.PresetGeometry>();
                shapeProperties.RemoveAllChildren<A.SolidFill>();
                shapeProperties.RemoveAllChildren<A.GradientFill>();

                var gradientFill = new A.GradientFill(
                    new A.GradientStopList(
                        new A.GradientStop(new A.SchemeColor(new A.Alpha() { Val = 45000 }) { Val = A.SchemeColorValues.Accent1 }) { Position = 0 },
                        new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 30000 }) { Val = A.SchemeColorValues.Accent2 }) { Position = 100000 }),
                    new A.LinearGradientFill() { Angle = 2700000 });
                if (geometry != null) {
                    shapeProperties.InsertAfter(gradientFill, geometry);
                } else {
                    shapeProperties.Append(gradientFill);
                }

                A.Outline outline = shapeProperties.GetFirstChild<A.Outline>() ?? new A.Outline();
                if (outline.Parent == null) {
                    shapeProperties.Append(outline);
                }

                outline.Width = 12700;
                outline.RemoveAllChildren();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() { Val = 65000 }) { Val = "336699" }));
                package.MainDocumentPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(docPath)) {
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before gradient shape", allText);
                Assert.Contains("After gradient shape", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("/ShadingType 2", content);
            Assert.Contains("/C0 [0.267 0.447 0.769] /C1 [0.949 0.643 0.435]", content);
            Assert.Contains("/Type /ExtGState /ca 0.45 /CA 0.65", content);
            Assert.Contains("0.2 0.4 0.6 RG", content);
            Assert.Contains("/SH1 sh", content);
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
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
