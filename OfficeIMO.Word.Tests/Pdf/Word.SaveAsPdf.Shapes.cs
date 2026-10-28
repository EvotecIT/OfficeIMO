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
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Simple_Shapes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeShapes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeShapes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native shapes");
                document.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(36), OfficeIMO.OfficeColor.Parse("#ccb399"), OfficeIMO.OfficeColor.Parse("#1a334d"), OfficeIMO.OfficeLength.Points(2.5));
                document.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeColor.Parse("#d0e6ff"), OfficeIMO.OfficeColor.Parse("#224466"), OfficeIMO.OfficeLength.Points(1.25));
                document.AddShape(WordShapeType.Line, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Parse("#ffffff"), OfficeIMO.OfficeColor.Parse("#008000"), OfficeIMO.OfficeLength.Points(2));
                document.AddParagraph("After native shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before native shapes", allText);
                Assert.Contains("After native shapes", allText);
            }

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
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
            WordShape line = document.AddParagraph().AddShapeDrawing(WordShapeType.Line, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(24));

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
                document.AddParagraph().AddShapeDrawing(WordShapeType.Line, OfficeIMO.OfficeLength.Points(96), OfficeIMO.OfficeLength.Points(1));
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
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
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
                WordShape triangle = document.AddParagraph().AddShapeDrawing(WordShapeType.Triangle, OfficeIMO.OfficeLength.Points(72), OfficeIMO.OfficeLength.Points(48));
                triangle.FillColorHex = "#99ccff";
                triangle.StrokeColorHex = "#003366";
                triangle.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);
                Assert.Equal(1.5, triangle.StrokeWeight!.Value.ToPoints());
                WordShape diamond = document.AddParagraph().AddShapeDrawing(WordShapeType.Diamond, OfficeIMO.OfficeLength.Points(64), OfficeIMO.OfficeLength.Points(48));
                diamond.FillColorHex = "#ffe699";
                WordShape arrow = document.AddParagraph().AddShapeDrawing(WordShapeType.RightArrow, OfficeIMO.OfficeLength.Points(96), OfficeIMO.OfficeLength.Points(36));
                arrow.FillColorHex = "#b7e1cd";
                WordShape star = document.AddParagraph().AddShapeDrawing(WordShapeType.Star5, OfficeIMO.OfficeLength.Points(56), OfficeIMO.OfficeLength.Points(56));
                star.FillColorHex = "#f4b183";
                document.AddParagraph("After native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before native DrawingML shapes", allText);
                Assert.Contains("After native DrawingML shapes", allText);
            }

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
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
                document.AddParagraph().AddShapeDrawing(WordShapeType.Heart, OfficeIMO.OfficeLength.Points(64), OfficeIMO.OfficeLength.Points(56)).FillColorHex = "#ff0000";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Cloud, OfficeIMO.OfficeLength.Points(88), OfficeIMO.OfficeLength.Points(52)).FillColorHex = "#00ff00";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Donut, OfficeIMO.OfficeLength.Points(64), OfficeIMO.OfficeLength.Points(64)).FillColorHex = "#0000ff";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Can, OfficeIMO.OfficeLength.Points(64), OfficeIMO.OfficeLength.Points(64)).FillColorHex = "#ffff00";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Cube, OfficeIMO.OfficeLength.Points(72), OfficeIMO.OfficeLength.Points(60)).FillColorHex = "#00ffff";
                document.AddParagraph("After remaining native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before remaining native DrawingML shapes", allText);
                Assert.Contains("After remaining native DrawingML shapes", allText);
            }

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
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
                document.AddParagraph().AddShapeDrawing(WordShapeType.Parallelogram, OfficeIMO.OfficeLength.Points(72), OfficeIMO.OfficeLength.Points(40)).FillColorHex = "#ff00ff";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Trapezoid, OfficeIMO.OfficeLength.Points(72), OfficeIMO.OfficeLength.Points(40)).FillColorHex = "#ff9900";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Chevron, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(44)).FillColorHex = "#808080";
                document.AddParagraph().AddShapeDrawing(WordShapeType.Plus, OfficeIMO.OfficeLength.Points(56), OfficeIMO.OfficeLength.Points(56)).FillColorHex = "#008080";
                document.AddParagraph().AddShapeDrawing(WordShapeType.LeftRightArrow, OfficeIMO.OfficeLength.Points(96), OfficeIMO.OfficeLength.Points(40)).FillColorHex = "#800000";
                document.AddParagraph("After additional native DrawingML shapes");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before additional native DrawingML shapes", allText);
                Assert.Contains("After additional native DrawingML shapes", allText);
            }

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
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
                document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(96), OfficeIMO.OfficeLength.Points(44));
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
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
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
                document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(112), OfficeIMO.OfficeLength.Points(48));
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
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before gradient shape", allText);
                Assert.Contains("After gradient shape", allText);
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            var (startRgb, endRgb) = ReadDrawingGradientEndpoints(bytes, "DeviceRGB", 3);
            Assert.Equal(OfficeColor.FromRgb(68, 114, 196), OfficeColor.FromRgb(
                (byte)System.Math.Round(startRgb[0] * 255D),
                (byte)System.Math.Round(startRgb[1] * 255D),
                (byte)System.Math.Round(startRgb[2] * 255D)));
            // DrawingML tint retains 30% of the input in linear-light sRGB;
            // the remaining 70% is white, matching native Office's tint contract.
            Assert.Equal(OfficeColor.FromRgb(250, 226, 219), OfficeColor.FromRgb(
                (byte)System.Math.Round(endRgb[0] * 255D),
                (byte)System.Math.Round(endRgb[1] * 255D),
                (byte)System.Math.Round(endRgb[2] * 255D)));

            var (startAlpha, endAlpha) = ReadDrawingGradientEndpoints(bytes, "DeviceGray", 1);
            // Drawing colors quantize alpha to a byte before PDF serialization.
            Assert.InRange(startAlpha[0], 0.45D - 1D / 255D, 0.45D + 1D / 255D);
            Assert.Equal(1D, endAlpha[0], precision: 6);

            var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetIdentityVisualPrimitives();
            var fill = Assert.Single(primitives, primitive => primitive.FillGradient != null || primitive.FunctionPaint != null);
            // A second global 45% opacity would fade the opaque endpoint and square
            // the translucent endpoint's alpha, even though the mask itself is right.
            Assert.Equal(1D, fill.FillOpacity ?? 1D, precision: 6);
            var stroke = Assert.Single(primitives, primitive => primitive.StrokeColor == OfficeColor.FromRgb(51, 102, 153));
            Assert.Equal(112D, stroke.Width, precision: 3);
            Assert.Equal(48D, stroke.Height, precision: 3);
            Assert.Equal(1D, stroke.StrokeWidth, precision: 3);
            Assert.InRange(stroke.StrokeOpacity ?? 1D, 0.65D - 1D / 255D, 0.65D + 1D / 255D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_DrawingML_Solid_Fill_Alpha_Once() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingSolidAlphaShape.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDrawingSolidAlphaShape.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before solid alpha shape");
                document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(112), OfficeIMO.OfficeLength.Points(48));
                document.AddParagraph("After solid alpha shape");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
                Wps.ShapeProperties properties = package.MainDocumentPart!.Document.Descendants<Wps.ShapeProperties>().First();
                A.PresetGeometry? geometry = properties.GetFirstChild<A.PresetGeometry>();
                properties.RemoveAllChildren<A.SolidFill>();
                properties.RemoveAllChildren<A.GradientFill>();
                var fill = new A.SolidFill(new A.SchemeColor(new A.Alpha() { Val = 45000 }) { Val = A.SchemeColorValues.Accent1 });
                if (geometry != null) properties.InsertAfter(fill, geometry);
                else properties.Append(fill);
                A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
                if (outline.Parent == null) properties.Append(outline);
                outline.Width = 12700;
                outline.RemoveAllChildren();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() { Val = 65000 }) { Val = "336699" }));
                package.MainDocumentPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(docPath)) {
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions { IncludePageNumbers = false });
            }
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Before solid alpha shape", allText);
                Assert.Contains("After solid alpha shape", allText);
            }

            var primitives = PdfCore.PdfReadDocument.Open(File.ReadAllBytes(pdfPath)).Pages[0].GetIdentityVisualPrimitives();
            var solid = Assert.Single(primitives, primitive => primitive.FillColor == OfficeColor.FromRgb(68, 114, 196));
            // The PDF paint alpha must be 45%, not the ~20% obtained by applying
            // DrawingML's color alpha again as a global fill opacity.
            Assert.InRange(solid.FillOpacity ?? 1D, 0.45D - 1D / 255D, 0.45D + 1D / 255D);
            Assert.Equal(112D, solid.Width, precision: 3);
            Assert.Equal(48D, solid.Height, precision: 3);
            Assert.Equal(OfficeColor.FromRgb(51, 102, 153), solid.StrokeColor);
            Assert.Equal(1D, solid.StrokeWidth, precision: 3);
            Assert.InRange(solid.StrokeOpacity ?? 1D, 0.65D - 1D / 255D, 0.65D + 1D / 255D);
        }

        private static (double[] Start, double[] End) ReadDrawingGradientEndpoints(byte[] pdf, string colorSpace, int components) {
            var (objects, _) = PdfCore.PdfSyntax.ParseObjects(pdf);
            var shading = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfCore.PdfDictionary>(),
                dictionary => dictionary.Get<PdfCore.PdfNumber>("ShadingType")?.Value == 2D &&
                    dictionary.Get<PdfCore.PdfName>("ColorSpace")?.Name == colorSpace);
            Assert.True(PdfCore.PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
                shading.Items["Function"], components, objects, 1024 * 1024, out var function));
            var start = new double[components];
            var end = new double[components];
            Assert.True(function.TryEvaluate(new[] { 0D }, start));
            Assert.True(function.TryEvaluate(new[] { 1D }, end));
            return (start, end);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Simple_TextBoxes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTextBoxes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTextBoxes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native text box");
                WordTextBox textBox = document.AddTextBox("Native text box body");
                textBox.WidthCentimeters = 7;
                textBox.HorizontalAlignment = WordTextBoxHorizontalAlignment.Center;
                textBox.Paragraphs[0].Bold = true;
                document.AddParagraph("After native text box");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
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

            string content = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
            Assert.Contains(" re S", content);
        }
    }
}
