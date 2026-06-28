using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests {
    public partial class WordImageExportTests {
        [Fact]
        public void WordDocument_ExportsFirstPageToPngAndSvgThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            WordParagraph title = document.AddParagraph("Shared Word renderer");
            title.SetFontSize(18).SetFontFamily("Calibri").SetColor(OfficeColor.FromRgb(17, 34, 51)).SetBold().SetAlignment(JustificationValues.Center);
            document.AddParagraph("This body paragraph is projected through OfficeIMO.Drawing text primitives.");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { Scale = 2D, BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult scaledSvg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { Scale = 2D, BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            Assert.Equal((int)Math.Ceiling(snapshot.Width * 2D), png.Width);
            Assert.Equal((int)Math.Ceiling(snapshot.Height * 2D), png.Height);
            Assert.Equal((int)Math.Ceiling(snapshot.Width * 2D), scaledSvg.Width);
            Assert.Equal((int)Math.Ceiling(snapshot.Height * 2D), scaledSvg.Height);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Shared Word renderer" && drawingText.Alignment == OfficeTextAlignment.Center);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text.StartsWith("This body paragraph", StringComparison.Ordinal));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);
            Assert.Equal(OfficeColor.White, image.GetPixel(2, 2));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("width=\"595.3pt\"", svgText, StringComparison.Ordinal);
            Assert.Contains("Shared Word renderer", svgText, StringComparison.Ordinal);
            string scaledSvgText = Encoding.UTF8.GetString(scaledSvg.Bytes);
            Assert.Contains("width=\"1190.6pt\"", scaledSvgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ToImageFluentExportsFirstPagePng() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Fluent Word image export");

            OfficeImageExportResult png = document.ToImage()
                .FirstPage()
                .HighResolution()
                .ExportPng();

            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);
            Assert.Empty(png.Diagnostics);
        }

        [Fact]
        public void WordDocument_ToImageFriendlyAliasesExportPngAndSvg() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Friendly Word image export");

            byte[] png = document.ToImage()
                .FirstPage()
                .ToPng();
            string svg = document.ToImage()
                .FirstPage()
                .ToSvg();

            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Friendly Word image export", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsLineSpacingHundredthsWithoutDoubleCountingFontSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try {
                using (WordDocument createdDocument = WordDocument.Create(filePath)) {
                    createdDocument.PageSettings.PageSize = WordPageSize.A4;
                    createdDocument.Margins.Type = WordMargin.Narrow;
                    createdDocument.AddParagraph("Before lines").SetFontSize(12);
                    createdDocument.AddParagraph("After lines").SetFontSize(12);
                    createdDocument.Save();
                }

                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                    Paragraph paragraph = package.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
                    paragraph.ParagraphProperties ??= new ParagraphProperties();
                    paragraph.ParagraphProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    paragraph.ParagraphProperties.SpacingBetweenLines.AfterLines = 100;
                    package.MainDocumentPart.Document.Save();
                }

                using WordDocument loadedDocument = WordDocument.Load(filePath);
                WordDocumentVisualSnapshot snapshot = loadedDocument.CreateVisualSnapshot();

                OfficeDrawingText before = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Before lines");
                OfficeDrawingText after = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "After lines");
                double gap = after.Y - (before.Y + before.Height);

                Assert.InRange(gap, 14.0D, 16.0D);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void WordDocument_ProjectsInlineDrawingShapesThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before shape");
            WordShape wordShape = document.AddParagraph().AddShapeDrawing(ShapeType.RightArrow, 96D, 40D);
            wordShape.FillColor = OfficeColor.FromRgb(14, 165, 233);
            wordShape.StrokeColor = OfficeColor.FromRgb(15, 23, 42);
            wordShape.StrokeWeight = 2D;
            document.AddParagraph("After shape");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingShape arrow = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .First(shape => shape.Shape.FillColor == OfficeColor.FromRgb(14, 165, 233));
            Assert.Equal(OfficeShapeKind.Polygon, arrow.Shape.Kind);
            Assert.Equal(96D, arrow.Shape.Width, 1);
            Assert.Equal(40D, arrow.Shape.Height, 1);
            Assert.Equal(OfficeColor.FromRgb(15, 23, 42), arrow.Shape.StrokeColor);
            Assert.Equal(2D, arrow.Shape.StrokeWidth);
            OfficeDrawingText afterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .First(text => text.Text == "After shape");
            Assert.True(afterText.Y > arrow.Y + arrow.Shape.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(14, 165, 233)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#0EA5E9", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0F172A", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsSquareWrappedAnchoredDrawingShapesThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;
            WordShape wordShape = document.AddParagraph().AddShapeDrawing(ShapeType.RightArrow, 96D, 40D, 144D, 36D);
            wordShape.FillColor = OfficeColor.FromRgb(34, 197, 94);
            wordShape.StrokeColor = OfficeColor.FromRgb(21, 128, 61);
            wordShape.StrokeWeight = 2D;
            document.AddParagraph("After anchored shape wraps beside the marker.");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingShape arrow = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .First(shape => shape.Shape.FillColor == OfficeColor.FromRgb(34, 197, 94));
            Assert.Equal(144D, arrow.X, 1);
            Assert.Equal(36D, arrow.Y, 1);
            Assert.Equal(96D, arrow.Shape.Width, 1);
            Assert.Equal(40D, arrow.Shape.Height, 1);
            OfficeDrawingText afterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "After anchored shape wraps beside the marker.");
            Assert.True(afterText.X >= arrow.X + arrow.Shape.Width - 0.5D);
            Assert.True(afterText.Y >= arrow.Y);
            Assert.True(afterText.Y < arrow.Y + arrow.Shape.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(34, 197, 94)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#22C55E", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#15803D", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsShapeTransformsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            WordShape drawingShape = document.AddParagraph().AddShapeDrawing(ShapeType.RightArrow, 96D, 40D);
            drawingShape.FillColor = OfficeColor.FromRgb(99, 102, 241);
            drawingShape.StrokeColor = OfficeColor.FromRgb(49, 46, 129);
            drawingShape.StrokeWeight = 2D;
            A.Transform2D drawingTransform = drawingShape._wpsShape!
                .GetFirstChild<Wps.ShapeProperties>()!
                .GetFirstChild<A.Transform2D>()!;
            drawingTransform.Rotation = 900000;
            drawingTransform.HorizontalFlip = true;

            WordShape vmlShape = document.AddParagraph().AddShape(
                ShapeType.Rectangle,
                84D,
                30D,
                OfficeColor.FromRgb(251, 146, 60),
                OfficeColor.FromRgb(154, 52, 18),
                2D);
            vmlShape.Rotation = 12D;
            vmlShape._rectangle!.Style = vmlShape._rectangle.Style!.Value + ";flip:y";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingShape transformedDrawingShape = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(99, 102, 241));
            OfficeTransform drawingShapeTransform = transformedDrawingShape.Shape.Transform!.Value;
            Assert.NotEqual(OfficeTransform.Identity, drawingShapeTransform);
            Assert.True(drawingShapeTransform.M11 < 0D);
            Assert.True(drawingShapeTransform.M12 < 0D);

            OfficeDrawingShape transformedVmlShape = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(251, 146, 60));
            OfficeTransform vmlShapeTransform = transformedVmlShape.Shape.Transform!.Value;
            Assert.NotEqual(OfficeTransform.Identity, vmlShapeTransform);
            Assert.True(vmlShapeTransform.M11 > 0D);
            Assert.True(vmlShapeTransform.M22 < 0D);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(99, 102, 241)) > 20);
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(251, 146, 60)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("transform=", svgText, StringComparison.Ordinal);
            Assert.Contains("#6366F1", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FB923C", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsVmlShapesThroughSharedDrawingPrimitives() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            document.AddParagraph().AddShape(ShapeType.Rectangle, 84D, 30D, OfficeColor.FromRgb(248, 113, 113), OfficeColor.FromRgb(127, 29, 29), 2D);
            document.AddParagraph().AddShape(ShapeType.Ellipse, 72D, 32D, OfficeColor.FromRgb(96, 165, 250), OfficeColor.FromRgb(30, 64, 175), 1.5D);
            document.AddParagraph().AddShape(ShapeType.RoundedRectangle, 90D, 34D, OfficeColor.FromRgb(250, 204, 21), OfficeColor.FromRgb(161, 98, 7), 2D, 0.2D);
            document.AddParagraph().AddShape(ShapeType.Line, 70D, 18D, OfficeColor.Transparent, OfficeColor.FromRgb(22, 163, 74), 3D);
            WordShape.AddPolygon(document.AddParagraph(), "0,30 30,0 60,30", "#C084FC", "#6B21A8");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            IReadOnlyList<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.FillColor == OfficeColor.FromRgb(248, 113, 113));
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Ellipse && shape.Shape.FillColor == OfficeColor.FromRgb(96, 165, 250));
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.RoundedRectangle && shape.Shape.FillColor == OfficeColor.FromRgb(250, 204, 21) && shape.Shape.CornerRadius > 0D);
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == OfficeColor.FromRgb(22, 163, 74) && shape.Shape.StrokeWidth == 3D);
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Polygon && shape.Shape.FillColor == OfficeColor.FromRgb(192, 132, 252));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(248, 113, 113)) > 50);
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(96, 165, 250)) > 50);
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(250, 204, 21)) > 50);
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(192, 132, 252)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("<ellipse", svgText, StringComparison.Ordinal);
            Assert.Contains("<line", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#F87171", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#60A5FA", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FACC15", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#C084FC", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsTextBoxesThroughSharedDrawingPrimitives() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            document.AddParagraph("Before text boxes");
            document.AddParagraph().AddTextBox("DrawingML text box content", WrapTextImage.InLineWithText);
            document.AddParagraph().AddTextBoxVml("Legacy VML text box content");
            document.AddParagraph("After text boxes");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            IReadOnlyList<OfficeDrawingText> texts = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().ToList();
            Assert.Contains(texts, text => text.Text.Contains("DrawingML text box content", StringComparison.Ordinal));
            Assert.Contains(texts, text => text.Text.Contains("Legacy VML text box content", StringComparison.Ordinal));
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count(shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.StrokeColor == OfficeColor.Black &&
                shape.Shape.StrokeWidth > 0D) >= 2);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.Black) > 40);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("DrawingML", svgText, StringComparison.Ordinal);
            Assert.Contains("Legacy", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTextBoxVerticalAlignmentThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            WordTextBox drawingTextBox = document.AddParagraph().AddTextBox("DrawingML bottom text box", WrapTextImage.InLineWithText);
            drawingTextBox.TextBodyProperties.Anchor = A.TextAnchoringTypeValues.Bottom;

            document.AddParagraph().AddTextBoxVml("VML middle text box");
            V.Shape vmlShape = document.BodyRoot.Descendants<V.Shape>()
                .Last(item => item.Descendants<V.TextBox>().Any());
            vmlShape.Style = "width:120pt;height:48pt;mso-wrap-style:square;v-text-anchor:middle";
            vmlShape.FillColor = "#E0F2FE";
            vmlShape.StrokeColor = "#0369A1";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText drawingText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(item => item.Text.Contains("DrawingML bottom", StringComparison.Ordinal));
            OfficeDrawingText vmlText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(item => item.Text.Contains("VML middle", StringComparison.Ordinal));
            Assert.Equal(OfficeTextVerticalAlignment.Bottom, drawingText.VerticalAlignment);
            Assert.Equal(OfficeTextVerticalAlignment.Center, vmlText.VerticalAlignment);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(224, 242, 254)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("DrawingML bottom", svgText, StringComparison.Ordinal);
            Assert.Contains("VML middle", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTextBoxRichRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            WordTextBox drawingTextBox = document.AddParagraph().AddTextBox("placeholder", WrapTextImage.InLineWithText);
            TextBoxContent drawingContent = drawingTextBox.Content!;
            Paragraph drawingParagraph = drawingContent.GetFirstChild<Paragraph>()!;
            drawingParagraph.RemoveAllChildren<Run>();
            drawingParagraph.Append(
                CreateWordTextBoxRun("Box ", "111827"),
                CreateWordTextBoxRun("Red", "DC2626", bold: true),
                CreateWordTextBoxRun(" blue", "2563EB", italic: true, underline: true));

            WordTextBox vmlTextBox = document.AddParagraph().AddTextBoxVml("placeholder");
            TextBoxContent vmlContent = vmlTextBox.Content!;
            Paragraph vmlParagraph = vmlContent.GetFirstChild<Paragraph>()!;
            vmlParagraph.RemoveAllChildren<Run>();
            vmlParagraph.Append(
                CreateWordTextBoxRun("Legacy ", "475569"),
                CreateWordTextBoxRun("Green", "16A34A", bold: true));
            V.Shape vmlShape = document.BodyRoot.Descendants<V.Shape>()
                .Last(item => item.Descendants<V.TextBox>().Any());
            vmlShape.Style = "width:120pt;height:44pt;mso-wrap-style:square;v-text-anchor:top";
            vmlShape.FillColor = "#FEF3C7";
            vmlShape.StrokeColor = "#92400E";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            List<OfficeDrawingRichText> richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToList();
            OfficeDrawingRichText drawingRichText = richTexts.Single(item => item.PlainText == "Box Red blue");
            OfficeDrawingRichText vmlRichText = richTexts.Single(item => item.PlainText == "Legacy Green");
            Assert.Equal(3, drawingRichText.Runs.Count);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), drawingRichText.Runs[1].Color);
            Assert.True(drawingRichText.Runs[1].Bold);
            Assert.Equal(OfficeColor.FromRgb(37, 99, 235), drawingRichText.Runs[2].Color);
            Assert.True(drawingRichText.Runs[2].Italic);
            Assert.True(drawingRichText.Runs[2].Underline);
            Assert.Equal(2, vmlRichText.Runs.Count);
            Assert.Equal(OfficeColor.FromRgb(22, 163, 74), vmlRichText.Runs[1].Color);
            Assert.True(vmlRichText.Runs[1].Bold);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(254, 243, 199)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Box", svgText, StringComparison.Ordinal);
            Assert.Contains("Legacy", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsTextBoxTransformsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            WordTextBox drawingTextBox = document.AddParagraph().AddTextBox("Rotated Word text", WrapTextImage.InLineWithText);
            A.Transform2D drawingTransform = drawingTextBox.DrawingShapeProperties!.GetFirstChild<A.Transform2D>()!;
            drawingTransform.Rotation = 900000;
            drawingTransform.HorizontalFlip = true;

            WordTextBox richTextBox = document.AddParagraph().AddTextBox("placeholder", WrapTextImage.InLineWithText);
            Paragraph richParagraph = richTextBox.Content!.GetFirstChild<Paragraph>()!;
            richParagraph.RemoveAllChildren<Run>();
            richParagraph.Append(
                CreateWordTextBoxRun("Rich ", "111827"),
                CreateWordTextBoxRun("Turn", "DC2626", bold: true));
            A.Transform2D richTransform = richTextBox.DrawingShapeProperties!.GetFirstChild<A.Transform2D>()!;
            richTransform.Rotation = 600000;
            richTransform.VerticalFlip = true;

            document.AddParagraph().AddTextBoxVml("VML turned text");
            V.Shape vmlShape = document.BodyRoot.Descendants<V.Shape>()
                .Last(item => item.Descendants<V.TextBox>().Any());
            vmlShape.Style = "width:132pt;height:46pt;mso-wrap-style:square;rotation:12;flip:y";
            vmlShape.FillColor = "#EDE9FE";
            vmlShape.StrokeColor = "#6D28D9";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);

            OfficeDrawingText drawingText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(item => item.Text.Contains("Rotated Word text", StringComparison.Ordinal));
            Assert.Equal(15D, drawingText.RotationDegrees, 6);
            Assert.True(drawingText.FlipHorizontal);
            Assert.False(drawingText.FlipVertical);

            OfficeDrawingRichText richText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingRichText>()
                .Single(item => item.PlainText == "Rich Turn");
            Assert.Equal(10D, richText.RotationDegrees, 6);
            Assert.False(richText.FlipHorizontal);
            Assert.True(richText.FlipVertical);

            OfficeDrawingText vmlText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(item => item.Text.Contains("VML turned text", StringComparison.Ordinal));
            Assert.Equal(12D, vmlText.RotationDegrees, 6);
            Assert.False(vmlText.FlipHorizontal);
            Assert.True(vmlText.FlipVertical);
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count(shape => shape.Shape.Transform.HasValue) >= 3);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(237, 233, 254)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Rotated Word", svgText, StringComparison.Ordinal);
            Assert.Contains("Rich", svgText, StringComparison.Ordinal);
            Assert.Contains("VML turned", svgText, StringComparison.Ordinal);
            Assert.Contains("transform=", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsVmlTextBoxInsetsAndFitToTextThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;

            document.AddParagraph("Before fit text box");
            document.AddParagraph().AddTextBoxVml("Fit VML text box content wraps into a taller frame.");
            V.Shape shape = document.BodyRoot.Descendants<V.Shape>()
                .Last(item => item.Descendants<V.TextBox>().Any());
            shape.Style = "width:96pt;mso-wrap-style:square;mso-fit-shape-to-text:t";
            shape.FillColor = "#DCFCE7";
            shape.StrokeColor = "#166534";
            shape.Stroked = true;
            shape.StrokeWeight = "2pt";
            shape.Descendants<V.TextBox>().Single().Inset = "12pt,6pt,18pt,9pt";
            document.AddParagraph("After fit text box");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(item => item.Text.Contains("Fit VML text box content", StringComparison.Ordinal));
            OfficeDrawingShape frame = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(item => item.Shape.FillColor == OfficeColor.FromRgb(220, 252, 231));
            Assert.Equal(12D, text.Padding.Left, 1);
            Assert.Equal(6D, text.Padding.Top, 1);
            Assert.Equal(18D, text.Padding.Right, 1);
            Assert.Equal(9D, text.Padding.Bottom, 1);
            Assert.Equal(96D, frame.Shape.Width, 1);
            Assert.True(frame.Shape.Height > 32D);
            Assert.Equal(frame.Shape.Height, text.Height, 1);
            Assert.Equal(OfficeColor.FromRgb(22, 101, 52), frame.Shape.StrokeColor);
            Assert.Equal(2D, frame.Shape.StrokeWidth, 1);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(220, 252, 231)) > 50);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Fit", svgText, StringComparison.Ordinal);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#166534", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_SaveAsImageHelpersUseSharedBuilderOptions() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.PageSize = WordPageSize.A4;
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Saved Word content");

            var options = new WordImageExportOptions {
                BackgroundColor = OfficeColor.White,
                IncludeDocumentContent = false,
                Scale = 2D
            };
            string folder = Path.Combine(Path.GetTempPath(), "officeimo-word-save-images-" + Guid.NewGuid().ToString("N"));
            try {
                string svgPath = Path.Combine(folder, "first-page.svg");
                document.SaveAsSvg(svgPath, options);

                Assert.True(File.Exists(svgPath));
                string svgText = File.ReadAllText(svgPath);
                Assert.Contains("<svg", svgText, StringComparison.Ordinal);
                Assert.DoesNotContain("Saved Word content", svgText, StringComparison.Ordinal);

                using var output = new MemoryStream();
                document.SaveAsPng(output, options);
                Assert.True(OfficePngReader.TryDecode(output.ToArray(), out OfficeRasterImage? image));
                WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot(options);
                Assert.Equal((int)Math.Ceiling(snapshot.Width * 2D), image!.Width);
                Assert.Equal((int)Math.Ceiling(snapshot.Height * 2D), image.Height);
                Assert.Equal(OfficeColor.White, image.GetPixel(2, 2));
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void WordDocument_ProjectsTablesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "North";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "South";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "East";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "West";
            table.Rows[0].Cells[0].ShadingFillColor = OfficeColor.FromRgb(204, 238, 255);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-tables");
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-tables");
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count() >= 4);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "North");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "West");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            OfficeDrawingShape filledCell = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .First(shape => shape.Shape.FillColor == OfficeColor.FromRgb(204, 238, 255));
            Assert.Equal(OfficeColor.FromRgb(204, 238, 255), image!.GetPixel((int)(filledCell.X + (filledCell.Shape.Width / 2D)), (int)(filledCell.Y + (filledCell.Shape.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("North", svgText, StringComparison.Ordinal);
            Assert.Contains("West", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Bordered";
            cell.ShadingFillColor = OfficeColor.FromRgb(248, 250, 252);
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "dc2626";
            cell.Borders.TopSize = 16U;
            cell.Borders.RightStyle = BorderValues.Dashed;
            cell.Borders.RightColorHex = "16a34a";
            cell.Borders.RightSize = 24U;
            cell.Borders.BottomStyle = BorderValues.Dotted;
            cell.Borders.BottomColorHex = "2563eb";
            cell.Borders.BottomSize = 8U;
            cell.Borders.LeftStyle = BorderValues.DotDash;
            cell.Borders.LeftColorHex = "9333ea";
            cell.Borders.LeftSize = 32U;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(22, 163, 74) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(37, 99, 235) && line.Shape.StrokeWidth == 1D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dot);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(147, 51, 234) && line.Shape.StrokeWidth == 4D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.DashDot);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Bordered", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#9333EA", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void WordDocument_ProjectsDoubleTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Double";
            cell.Borders.TopStyle = BorderValues.Double;
            cell.Borders.TopColorHex = "dc2626";
            cell.Borders.TopSize = 16U;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            List<OfficeDrawingShape> redBorderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38))
                .ToList();
            Assert.True(redBorderLines.Count >= 2);
            Assert.All(redBorderLines, line => Assert.Equal(2D, line.Shape.StrokeWidth));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Double", svgText, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svgText, "#DC2626") >= 2);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void WordDocument_ProjectsDiagonalTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Diagonal";
            cell.Borders.TopLeftToBottomRightStyle = BorderValues.DotDash;
            cell.Borders.TopLeftToBottomRightColorHex = "dc2626";
            cell.Borders.TopLeftToBottomRightSize = 16U;
            cell.Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
            cell.Borders.TopRightToBottomLeftColorHex = "2563eb";
            cell.Borders.TopRightToBottomLeftSize = 24U;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            List<OfficeDrawingShape> diagonalLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                    (shape.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) || shape.Shape.StrokeColor == OfficeColor.FromRgb(37, 99, 235)))
                .ToList();
            Assert.Contains(diagonalLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 2D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.DashDot);
            Assert.Contains(diagonalLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(37, 99, 235) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dot);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Diagonal", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void WordDocument_ProjectsThemeTableCellFillAndBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");
            SetThemeColor(document, "accent2", "654321");
            SetThemeColor(document, "accent3", "ABCDEF");

            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Theme cell";
            cell.ShadingFillColorHex = "ffffff";
            cell._tableCellProperties!.Shading!.ThemeFill = ThemeColorValues.Accent1;
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "ffffff";
            cell.Borders.TopSize = 16U;
            cell._tableCellProperties!.TableCellBorders!.TopBorder!.ThemeColor = ThemeColorValues.Accent2;
            cell.Borders.TopLeftToBottomRightStyle = BorderValues.Single;
            cell.Borders.TopLeftToBottomRightColorHex = "ffffff";
            cell.Borders.TopLeftToBottomRightSize = 24U;
            cell._tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder!.ThemeColor = ThemeColorValues.Accent3;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape shape && shape.Shape.FillColor == OfficeColor.FromRgb(18, 52, 86));
            List<OfficeDrawingShape> themeLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(themeLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(101, 67, 33) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(themeLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(171, 205, 239) && line.Shape.StrokeWidth == 3D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme cell", svgText, StringComparison.Ordinal);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#654321", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#ABCDEF", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            OfficeDrawingShape fill = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .First(shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 52, 86));
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), image!.GetPixel((int)(fill.X + fill.Shape.Width - 6D), (int)(fill.Y + 6D)));
        }

        [Fact]
        public void WordDocument_ProjectsThemeTableCellTextThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent3", "ABCDEF");

            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Theme table text";
            paragraph.ThemeColor = ThemeColorValues.Accent3;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(element => element.Text == "Theme table text");
            Assert.Equal(OfficeColor.FromRgb(171, 205, 239), text.Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme table text", svgText, StringComparison.Ordinal);
            Assert.Contains("#ABCDEF", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsStyleInheritedTableCellFillAndBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");

            const string baseStyleId = "ImageTableBaseStyle";
            const string derivedStyleId = "ImageTableDerivedStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Table Base Style" },
                    new StyleTableProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "E0F2FE" },
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "ffffff", ThemeColor = ThemeColorValues.Accent1, Size = 16U },
                            new LeftBorder { Val = BorderValues.Dashed, Color = "dc2626", Size = 24U }))) {
                    Type = StyleValues.Table,
                    StyleId = baseStyleId,
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Table Derived Style" },
                    new BasedOn { Val = baseStyleId }) {
                    Type = StyleValues.Table,
                    StyleId = derivedStyleId,
                    CustomStyle = true
                });

            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            table._tableProperties!.TableStyle = new TableStyle { Val = derivedStyleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Style table cell";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(224, 242, 254));
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(18, 52, 86) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Style table cell", svgText, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsConditionalTableStyleFillsAndBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");

            const string tableStyleId = "ImageConditionalTableStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Conditional Table Style" },
                    new StyleTableProperties(
                        new TableStyleRowBandSize { Val = 2 },
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "94A3B8", Size = 8U },
                            new BottomBorder { Val = BorderValues.Single, Color = "94A3B8", Size = 8U },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Color = "CBD5E1", Size = 8U })),
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new TableCellBorders(
                                new BottomBorder { Val = BorderValues.Dashed, Color = "DC2626", Size = 24U }),
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "FFFFFF", ThemeFill = ThemeColorValues.Accent1 })) {
                        Type = TableStyleOverrideValues.FirstRow
                    },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "E2EFD9" })) {
                        Type = TableStyleOverrideValues.Band1Horizontal
                    },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "FEE2E2" })) {
                        Type = TableStyleOverrideValues.Band2Horizontal
                    }) {
                    Type = StyleValues.Table,
                    StyleId = tableStyleId,
                    CustomStyle = true
                });

            WordTable table = document.AddTable(5, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            table._tableProperties!.TableStyle = new TableStyle { Val = tableStyleId };
            table.ConditionalFormattingFirstRow = true;
            table.ConditionalFormattingNoHorizontalBand = false;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Conditional header";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Band one first";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "Band one second";
            table.Rows[3].Cells[0].Paragraphs[0].Text = "Band two first";
            table.Rows[4].Cells[0].Paragraphs[0].Text = "Band two second";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(18, 52, 86));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(226, 239, 217));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(254, 226, 226));
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Conditional header", svgText, StringComparison.Ordinal);
            Assert.Contains("Band one second", svgText, StringComparison.Ordinal);
            Assert.Contains("Band two first", svgText, StringComparison.Ordinal);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E2EFD9", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEE2E2", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsConditionalTableStyleColumnBandsThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;

            const string tableStyleId = "ImageColumnBandTableStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Column Band Table Style" },
                    new StyleTableProperties(new TableStyleColumnBandSize { Val = 2 }),
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "E0F2FE" })) {
                        Type = TableStyleOverrideValues.Band1Vertical
                    },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "FEF3C7" })) {
                        Type = TableStyleOverrideValues.Band2Vertical
                    }) {
                    Type = StyleValues.Table,
                    StyleId = tableStyleId,
                    CustomStyle = true
                });

            WordTable table = document.AddTable(1, 5);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 5000;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 1000, 1000, 1000, 1000, 1000 };
            table._tableProperties!.TableStyle = new TableStyle { Val = tableStyleId };
            table.ConditionalFormattingFirstRow = false;
            table.ConditionalFormattingFirstColumn = false;
            table.ConditionalFormattingNoHorizontalBand = true;
            table.ConditionalFormattingNoVerticalBand = false;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Band1A";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Band1B";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "Band2A";
            table.Rows[0].Cells[3].Paragraphs[0].Text = "Band2B";
            table.Rows[0].Cells[4].Paragraphs[0].Text = "Band1C";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(224, 242, 254));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(254, 243, 199));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Band1B");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Band2A");
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsConditionalTableStyleCornersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;

            const string tableStyleId = "ImageCornerTableStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Corner Table Style" },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "FEF3C7" })) {
                        Type = TableStyleOverrideValues.LastRow
                    },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "DBEAFE" })) {
                        Type = TableStyleOverrideValues.LastColumn
                    },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "FCE7F3" })) {
                        Type = TableStyleOverrideValues.SouthEastCell
                    }) {
                    Type = StyleValues.Table,
                    StyleId = tableStyleId,
                    CustomStyle = true
                });

            WordTable table = document.AddTable(2, 2);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 2100, 2100 };
            table._tableProperties!.TableStyle = new TableStyle { Val = tableStyleId };
            table.ConditionalFormattingLastRow = true;
            table.ConditionalFormattingLastColumn = true;
            table.ConditionalFormattingNoHorizontalBand = true;
            table.ConditionalFormattingNoVerticalBand = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Plain";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Last column";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Last row";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "South east corner";

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(254, 243, 199));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(219, 234, 254));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(252, 231, 243));
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingText text &&
                text.Text == "South east corner");

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellMarginsThroughSharedDrawingPadding() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Padded cell";
            cell.MarginLeftWidth = 360;
            cell.MarginTopWidth = 120;
            cell.MarginRightWidth = 240;
            cell.MarginBottomWidth = 80;
            cell.ShadingFillColor = OfficeColor.FromRgb(226, 239, 218);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);

            OfficeDrawingShape cellFrame = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(226, 239, 218));
            OfficeDrawingText text = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(drawingText => drawingText.Text == "Padded cell");

            Assert.Equal(cellFrame.X, text.X);
            Assert.Equal(cellFrame.Y, text.Y);
            Assert.Equal(cellFrame.Shape.Width, text.Width);
            Assert.Equal(cellFrame.Shape.Height, text.Height);
            Assert.True(text.HasPadding);
            Assert.Equal(18D, text.Padding.Left);
            Assert.Equal(6D, text.Padding.Top);
            Assert.Equal(12D, text.Padding.Right);
            Assert.Equal(4D, text.Padding.Bottom);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Padded cell", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellRichRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = "Plain ";
            paragraph.SetFontFamily("Aptos").SetFontSize(11).SetColor(OfficeColor.FromRgb(17, 24, 39));
            paragraph.AddText("Red").SetColor(OfficeColor.FromRgb(220, 38, 38)).SetBold();
            paragraph.AddText(" italic").SetColor(OfficeColor.FromRgb(37, 99, 235)).SetItalic().SetUnderline(UnderlineValues.Single);
            cell.MarginLeftWidth = 300;
            cell.MarginTopWidth = 100;
            cell.MarginRightWidth = 200;
            cell.MarginBottomWidth = 80;

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Plain Red italic", richText.PlainText);
            Assert.Equal(3, richText.Runs.Count);
            Assert.True(richText.HasPadding);
            Assert.Equal(15D, richText.Padding.Left);
            Assert.Equal(5D, richText.Padding.Top);
            Assert.Equal(10D, richText.Padding.Right);
            Assert.Equal(4D, richText.Padding.Bottom);
            Assert.Equal(OfficeColor.FromRgb(17, 24, 39), richText.Runs[0].Color);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), richText.Runs[1].Color);
            Assert.True(richText.Runs[1].Bold);
            Assert.Equal(OfficeColor.FromRgb(37, 99, 235), richText.Runs[2].Color);
            Assert.True(richText.Runs[2].Italic);
            Assert.True(richText.Runs[2].Underline);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Plain", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("italic", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellRunHighlightThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = "Marked cell";
            paragraph.SetFontSize(14);
            paragraph.SetHighlight(HighlightColorValues.Cyan);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.Equal("Marked cell", run.Text);
            Assert.Equal(OfficeColor.Cyan, run.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("#00FFFF", svgText, StringComparison.OrdinalIgnoreCase);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.Cyan) > 20, "Expected highlighted Word table-cell run background to render through the shared raster rich-text path.");
        }

        [Fact]
        public void WordDocument_ProjectsNestedTablesInsideTableCellsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable outer = document.AddTable(1, 1);
            outer.WidthType = TableWidthUnitValues.Dxa;
            outer.Width = 7200;
            outer.ColumnWidthType = TableWidthUnitValues.Dxa;
            outer.ColumnWidth = new List<int> { 7200 };
            WordTableCell hostCell = outer.Rows[0].Cells[0];
            hostCell.ShadingFillColor = OfficeColor.FromRgb(243, 244, 246);
            WordTable nested = hostCell.AddTable(1, 2, removePrecedingParagraph: true);
            nested.WidthType = TableWidthUnitValues.Dxa;
            nested.Width = 6000;
            nested.ColumnWidthType = TableWidthUnitValues.Dxa;
            nested.ColumnWidth = new List<int> { 3000, 3000 };
            nested.Rows[0].Cells[0].Paragraphs[0].Text = "Inner A";
            nested.Rows[0].Cells[1].Paragraphs[0].Text = "Inner B";
            nested.Rows[0].Cells[0].ShadingFillColor = OfficeColor.FromRgb(219, 234, 254);
            nested.Rows[0].Cells[1].ShadingFillColor = OfficeColor.FromRgb(220, 252, 231);
            hostCell.AddParagraph().AddText("After nested");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Inner A");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Inner B");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "After nested");
            OfficeDrawingText innerAText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Inner A");
            Assert.True(
                innerAText.Width > innerAText.Padding.Horizontal && innerAText.Height > innerAText.Padding.Vertical,
                $"Nested table text content rectangle must be positive. Size={innerAText.Width}x{innerAText.Height}; padding={innerAText.Padding.Horizontal}x{innerAText.Padding.Vertical}.");

            OfficeDrawingShape outerFrame = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(243, 244, 246));
            OfficeDrawingShape innerAFrame = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(219, 234, 254));
            OfficeDrawingShape innerBFrame = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(220, 252, 231));
            Assert.True(innerAFrame.X > outerFrame.X);
            Assert.True(innerBFrame.X > innerAFrame.X);
            Assert.True(innerAFrame.Y > outerFrame.Y);
            OfficeDrawingText afterNested = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "After nested");
            double nestedBottom = Math.Max(innerAFrame.Y + innerAFrame.Shape.Height, innerBFrame.Y + innerBFrame.Shape.Height);
            Assert.True(afterNested.Y >= nestedBottom, $"Expected following cell text below nested table, got text Y {afterNested.Y} and nested bottom {nestedBottom}.");

            string snapshotSvgText = OfficeDrawingSvgExporter.ToSvg(snapshot.Drawing);
            Assert.Contains("Inner A", snapshotSvgText, StringComparison.Ordinal);
            Assert.Contains("Inner B", snapshotSvgText, StringComparison.Ordinal);
            Assert.Contains("After nested", snapshotSvgText, StringComparison.Ordinal);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Inner A", svgText, StringComparison.Ordinal);
            Assert.Contains("Inner B", svgText, StringComparison.Ordinal);
            Assert.Contains("After nested", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void WordDocument_ProjectsListMarkersThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordList bullets = document.AddList(WordListStyle.Bulleted);
            bullets.AddItem("Bullet item");
            WordList numbered = document.AddList(WordListStyle.Numbered);
            numbered.AddItem("Numbered item");

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "1.");

            OfficeDrawingText bulletBody = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Bullet item");
            OfficeDrawingText bulletMarker = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Y == bulletBody.Y && text.X < bulletBody.X);
            Assert.True(bulletMarker.X < bulletBody.X);
            Assert.False(string.IsNullOrWhiteSpace(bulletMarker.Text));
            Assert.Equal("Symbol", bulletMarker.Font.FamilyName);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains(bulletMarker.Text, svgText, StringComparison.Ordinal);
            Assert.Contains("1.", svgText, StringComparison.Ordinal);
            Assert.Contains("Bullet item", svgText, StringComparison.Ordinal);
            Assert.Contains("Numbered item", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsRichRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordParagraph paragraph = document.AddParagraph("Plain ");
            paragraph.SetFontFamily("Aptos").SetFontSize(11).SetColor(OfficeColor.FromRgb(17, 24, 39));
            paragraph.AddText("Red").SetColor(OfficeColor.FromRgb(220, 38, 38)).SetBold();
            paragraph.AddText(" italic").SetColor(OfficeColor.FromRgb(37, 99, 235)).SetItalic().SetUnderline(UnderlineValues.Single);

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Plain Red italic", richText.PlainText);
            Assert.Equal(3, richText.Runs.Count);
            Assert.Equal(OfficeColor.FromRgb(17, 24, 39), richText.Runs[0].Color);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), richText.Runs[1].Color);
            Assert.True(richText.Runs[1].Bold);
            Assert.Equal(OfficeColor.FromRgb(37, 99, 235), richText.Runs[2].Color);
            Assert.True(richText.Runs[2].Italic);
            Assert.True(richText.Runs[2].Underline);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Plain", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("italic", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsThemeParagraphTextAndRichRunsThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");
            SetThemeColor(document, "accent2", "654321");
            SetThemeColor(document, "accent3", "ABCDEF");

            WordParagraph plain = document.AddParagraph("Theme body");
            plain.ThemeColor = ThemeColorValues.Accent1;
            WordParagraph rich = document.AddParagraph("Theme ");
            rich.ThemeColor = ThemeColorValues.Accent2;
            rich.AddText("rich").ThemeColor = ThemeColorValues.Accent3;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText bodyText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Theme body");
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), bodyText.Color);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Theme rich", richText.PlainText);
            Assert.Equal(OfficeColor.FromRgb(101, 67, 33), richText.Runs[0].Color);
            Assert.Equal(OfficeColor.FromRgb(171, 205, 239), richText.Runs[1].Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme body", svgText, StringComparison.Ordinal);
            Assert.Contains("Theme", svgText, StringComparison.Ordinal);
            Assert.Contains("rich", svgText, StringComparison.Ordinal);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#654321", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#ABCDEF", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsRunHighlightThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordParagraph paragraph = document.AddParagraph("Marked");
            paragraph.SetHighlight(HighlightColorValues.Yellow);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.Equal("Marked", run.Text);
            Assert.Equal(OfficeColor.Yellow, run.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFFF00", svgText, StringComparison.OrdinalIgnoreCase);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.Yellow) > 20, "Expected highlighted Word run background to render through the shared raster rich-text path.");
        }

        [Fact]
        public void WordDocument_ProjectsStyleInheritedRunHighlightsThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;

            const string baseCharacterStyleId = "ImageHighlightBaseChar";
            const string derivedCharacterStyleId = "ImageHighlightDerivedChar";
            const string baseParagraphStyleId = "ImageHighlightBaseParagraph";
            const string derivedParagraphStyleId = "ImageHighlightDerivedParagraph";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Highlight Base Character" },
                    new StyleRunProperties(new Highlight { Val = HighlightColorValues.Yellow })) {
                    Type = StyleValues.Character,
                    StyleId = baseCharacterStyleId,
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Highlight Derived Character" },
                    new BasedOn { Val = baseCharacterStyleId }) {
                    Type = StyleValues.Character,
                    StyleId = derivedCharacterStyleId,
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Highlight Base Paragraph" },
                    new StyleRunProperties(new Highlight { Val = HighlightColorValues.Cyan })) {
                    Type = StyleValues.Paragraph,
                    StyleId = baseParagraphStyleId,
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Highlight Derived Paragraph" },
                    new BasedOn { Val = baseParagraphStyleId }) {
                    Type = StyleValues.Paragraph,
                    StyleId = derivedParagraphStyleId,
                    CustomStyle = true
                });

            document.AddParagraph("Character style marked").SetCharacterStyleId(derivedCharacterStyleId);
            document.AddParagraph("Paragraph style marked").SetStyleId(derivedParagraphStyleId);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            List<OfficeDrawingRichText> richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToList();
            OfficeRichTextRun characterRun = Assert.Single(richTexts.Single(text => text.PlainText == "Character style marked").Runs);
            OfficeRichTextRun paragraphRun = Assert.Single(richTexts.Single(text => text.PlainText == "Paragraph style marked").Runs);
            Assert.Equal(OfficeColor.Yellow, characterRun.BackgroundColor);
            Assert.Equal(OfficeColor.Cyan, paragraphRun.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Character style marked", svgText, StringComparison.Ordinal);
            Assert.Contains("Paragraph style marked", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFFF00", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#00FFFF", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.Yellow) > 20, "Expected character-style inherited Word highlight to render through the shared raster rich-text path.");
            Assert.True(CountPixelsNear(image!, OfficeColor.Cyan) > 20, "Expected paragraph-style inherited Word highlight to render through the shared raster rich-text path.");
        }

        [Fact]
        public void WordDocument_ProjectsParagraphShadingAndBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");

            WordParagraph paragraph = document.AddParagraph("Framed paragraph");
            paragraph.ShadingFillColor = OfficeColor.FromRgb(226, 239, 218);
            paragraph.Borders.TopStyle = BorderValues.Single;
            paragraph.Borders.TopColorHex = "ffffff";
            paragraph.Borders.TopThemeColor = ThemeColorValues.Accent1;
            paragraph.Borders.TopSize = 16U;
            paragraph.Borders.LeftStyle = BorderValues.DotDash;
            paragraph.Borders.LeftColorHex = "dc2626";
            paragraph.Borders.LeftSize = 24U;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(element => element.Text == "Framed paragraph");
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(226, 239, 218) &&
                shape.X <= text.X &&
                shape.Y <= text.Y);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(18, 52, 86) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.DashDot);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Framed paragraph", svgText, StringComparison.Ordinal);
            Assert.Contains("#E2EFDA", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsStyleInheritedParagraphShadingAndBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            SetThemeColor(document, "accent1", "123456");

            const string baseStyleId = "ImageFrameBaseStyle";
            const string derivedStyleId = "ImageFrameDerivedStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Frame Base Style" },
                    new StyleParagraphProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "E0F2FE" },
                        new ParagraphBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "ffffff", ThemeColor = ThemeColorValues.Accent1, Size = 16U },
                            new LeftBorder { Val = BorderValues.Dashed, Color = "dc2626", Size = 24U }))) {
                    Type = StyleValues.Paragraph,
                    StyleId = baseStyleId,
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Frame Derived Style" },
                    new BasedOn { Val = baseStyleId }) {
                    Type = StyleValues.Paragraph,
                    StyleId = derivedStyleId,
                    CustomStyle = true
                });

            WordParagraph paragraph = document.AddParagraph("Style framed paragraph");
            paragraph.SetStyleId(derivedStyleId);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(element => element.Text == "Style framed paragraph");
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(224, 242, 254) &&
                shape.X <= text.X &&
                shape.Y <= text.Y);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(18, 52, 86) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Style framed paragraph", svgText, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHexBackedParagraphColorsAsOpaqueThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;

            WordParagraph paragraph = document.AddParagraph("Opaque Word paragraph");
            paragraph.SetColor(OfficeColor.FromRgba(17, 24, 39, 128));
            paragraph.ShadingFillColor = OfficeColor.FromRgba(226, 239, 218, 128);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(element => element.Text == "Opaque Word paragraph");
            Assert.Equal(OfficeColor.FromRgb(17, 24, 39), text.Color);
            Assert.Contains(snapshot.Drawing.Elements, element =>
                element is OfficeDrawingShape shape &&
                shape.Shape.FillColor == OfficeColor.FromRgb(226, 239, 218) &&
                shape.X <= text.X &&
                shape.Y <= text.Y);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Opaque Word paragraph", svgText, StringComparison.Ordinal);
            Assert.Contains("fill=\"#111827\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill=\"#E2EFDA\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("fill-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsParagraphIndentsThroughSharedDrawingPadding() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordParagraph plain = document.AddParagraph("Plain indented");
            plain.IndentationBeforePoints = 24D;
            plain.IndentationAfterPoints = 12D;

            WordParagraph rich = document.AddParagraph("Rich ");
            rich.IndentationBeforePoints = 18D;
            rich.IndentationAfterPoints = 6D;
            rich.AddText("indented").SetBold();

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);

            OfficeDrawingText plainText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Plain indented");
            Assert.True(plainText.HasPadding);
            Assert.Equal(24D, plainText.Padding.Left);
            Assert.Equal(12D, plainText.Padding.Right);

            OfficeDrawingRichText richText = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().Single(text => text.PlainText == "Rich indented");
            Assert.True(richText.HasPadding);
            Assert.Equal(18D, richText.Padding.Left);
            Assert.Equal(6D, richText.Padding.Right);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Plain indented", svgText, StringComparison.Ordinal);
            Assert.Contains("Rich", svgText, StringComparison.Ordinal);
            Assert.Contains("indented", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsFirstLineAndHangingIndentsThroughSharedDrawingTextIndent() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordParagraph hanging = document.AddParagraph("Hanging paragraph wraps across exported image lines");
            hanging.IndentationBeforePoints = 24D;
            hanging.IndentationHangingPoints = 12D;

            WordParagraph firstLine = document.AddParagraph("First line ");
            firstLine.IndentationBeforePoints = 10D;
            firstLine.IndentationFirstLinePoints = 16D;
            firstLine.AddText("rich indent").SetBold();

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);

            OfficeDrawingText hangingText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text.StartsWith("Hanging", StringComparison.Ordinal));
            Assert.Equal(12D, hangingText.Padding.Left);
            Assert.True(hangingText.HasParagraphIndent);
            Assert.Equal(0D, hangingText.ParagraphIndent.FirstLineOffset);
            Assert.Equal(12D, hangingText.ParagraphIndent.ContinuationLineOffset);

            OfficeDrawingRichText richText = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().Single(text => text.PlainText == "First line rich indent");
            Assert.Equal(10D, richText.Padding.Left);
            Assert.True(richText.HasParagraphIndent);
            Assert.Equal(16D, richText.ParagraphIndent.FirstLineOffset);
            Assert.Equal(0D, richText.ParagraphIndent.ContinuationLineOffset);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Hanging", svgText, StringComparison.Ordinal);
            Assert.Contains("rich indent", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsJustifiedParagraphsThroughSharedDrawingTextAlignment() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordParagraph paragraph = document.AddParagraph("Justified Word paragraph wraps across the exported preview and distributes text through the shared renderer.");
            paragraph.SetFontSize(12).SetAlignment(JustificationValues.Both);

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(element => element.Text.StartsWith("Justified Word", StringComparison.Ordinal));
            Assert.Equal(OfficeTextAlignment.Justify, text.Alignment);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Justified Word", svgText, StringComparison.Ordinal);
            Assert.Contains("textLength=", svgText, StringComparison.Ordinal);
            Assert.Contains("lengthAdjust=\"spacing\"", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ExpandsTabsInParagraphsThroughSharedDrawingLayout() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("A\tB").SetFontSize(12);

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            Assert.Empty(svg.Diagnostics);
            Assert.DoesNotContain("\t", svgText, StringComparison.Ordinal);
            Assert.Contains("xml:space=\"preserve\"", svgText, StringComparison.Ordinal);
            Assert.Contains("A   B", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsEmbeddedPngImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourcePng = CreateSolidPng(24, 18, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            document.AddParagraph().AddImage(imageStream, "inline.png", 24, 18, description: "Inline blue marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Inline blue marker", drawingImage.AlternativeText);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_RendersEmbeddedBmpImagesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourceBmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86),
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86)
            });
            using var imageStream = new MemoryStream(sourceBmp);
            document.AddParagraph().AddImage(imageStream, "inline.bmp", 24, 18, description: "Inline BMP marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Inline BMP marker", drawingImage.AlternativeText);
            Assert.Equal("image/bmp", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(18, 52, 86),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_RendersEmbeddedTopDownBmpImagesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourceBmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144),
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144)
            }, topDown: true);
            using var imageStream = new MemoryStream(sourceBmp);
            document.AddParagraph().AddImage(imageStream, "inline-top-down.bmp", 24, 18, description: "Inline top-down BMP marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Inline top-down BMP marker", drawingImage.AlternativeText);
            Assert.Equal("image/bmp", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(24, 96, 144),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_RendersEmbeddedBmp32AlphaImagesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourceBmp = CreateBmp32(2, 2, new[] {
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128),
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128)
            });
            using var imageStream = new MemoryStream(sourceBmp);
            document.AddParagraph().AddImage(imageStream, "inline-alpha.bmp", 24, 18, description: "Inline BMP32 alpha marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Inline BMP32 alpha marker", drawingImage.AlternativeText);
            Assert.Equal("image/bmp", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            OfficeColor blended = rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D)));
            Assert.True(blended.R >= 252, $"Expected red channel to stay near full after BMP alpha blend, got {blended.R}.");
            Assert.InRange(blended.G, 124, 130);
            Assert.InRange(blended.B, 124, 130);
            Assert.Equal(255, blended.A);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_RendersEmbeddedGifImagesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            using var imageStream = new MemoryStream(CreateSinglePixelGif());
            document.AddParagraph().AddImage(imageStream, "inline.gif", 24, 18, description: "Inline GIF marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.Black });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.Black });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Inline GIF marker", drawingImage.AlternativeText);
            Assert.Equal("image/gif", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.White,
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/gif;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsImageCropThroughSharedDrawingProjection() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourcePng = CreateSplitPng(40, 20, OfficeColor.FromRgb(220, 38, 38), OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage image = document.AddParagraph().InsertImage(imageStream, "cropped.png", 40, 20, description: "Cropped blue marker");
            image.CropLeft = 50000;

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal(0.5D, drawingImage.Projection.SourceLeft, 3);
            Assert.Equal(0.5D, drawingImage.Projection.SourceWidth, 3);
            Assert.True(drawingImage.Projection.HasCrop);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<clipPath", svgText, StringComparison.Ordinal);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsFirstPageHeaderAndFooterThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.HeaderDefaultOrCreate.AddParagraph("Default header");
            document.FooterDefaultOrCreate.AddParagraph("Default footer");
            document.DifferentFirstPage = true;
            document.HeaderFirstOrCreate.AddParagraph("First page header").SetAlignment(JustificationValues.Center);
            document.FooterFirstOrCreate.AddParagraph("First page footer").SetAlignment(JustificationValues.Center);
            document.AddParagraph("Body starts below header");

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText header = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "First page header");
            OfficeDrawingText footer = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "First page footer");
            OfficeDrawingText body = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Body starts below header");
            Assert.True(header.Y < body.Y);
            Assert.True(footer.Y > body.Y);
            Assert.Equal(OfficeTextAlignment.Center, header.Alignment);
            Assert.Equal(OfficeTextAlignment.Center, footer.Alignment);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("First page header", svgText, StringComparison.Ordinal);
            Assert.Contains("First page footer", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("Default header", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("Default footer", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_DoesNotProjectDefaultHeaderFooterWhenFirstPagePartsAreMissing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.HeaderDefaultOrCreate.AddParagraph("Default header");
            document.FooterDefaultOrCreate.AddParagraph("Default footer");
            document.DifferentFirstPage = true;
            document.AddParagraph("First page body");

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Default header");
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Default footer");
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "First page body");

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.DoesNotContain("Default header", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("Default footer", svgText, StringComparison.Ordinal);
            Assert.Contains("First page body", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterRichRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;
            WordParagraph header = document.HeaderFirstOrCreate.AddParagraph("Head ");
            header.SetFontFamily("Aptos").SetFontSize(11).SetColor(OfficeColor.FromRgb(17, 24, 39)).SetAlignment(JustificationValues.Center);
            header.AddText("Red").SetColor(OfficeColor.FromRgb(220, 38, 38)).SetBold();
            header.AddText(" blue").SetColor(OfficeColor.FromRgb(37, 99, 235)).SetItalic().SetUnderline(UnderlineValues.Single);
            WordParagraph footer = document.FooterFirstOrCreate.AddParagraph("Foot ");
            footer.SetFontFamily("Aptos").SetFontSize(10).SetColor(OfficeColor.FromRgb(55, 65, 81)).SetAlignment(JustificationValues.Center);
            footer.AddText("Green").SetColor(OfficeColor.FromRgb(22, 163, 74)).SetBold();
            document.AddParagraph("Body starts below rich header");

            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(svg.Diagnostics);
            Assert.Empty(snapshot.Diagnostics);
            var richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToList();
            OfficeDrawingRichText headerRichText = richTexts.Single(text => text.PlainText == "Head Red blue");
            OfficeDrawingRichText footerRichText = richTexts.Single(text => text.PlainText == "Foot Green");
            OfficeDrawingText body = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Body starts below rich header");

            Assert.True(headerRichText.Y < body.Y);
            Assert.True(footerRichText.Y > body.Y);
            Assert.Equal(OfficeTextAlignment.Center, headerRichText.Alignment);
            Assert.Equal(OfficeTextAlignment.Center, footerRichText.Alignment);
            Assert.Equal(3, headerRichText.Runs.Count);
            Assert.Equal(2, footerRichText.Runs.Count);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), headerRichText.Runs[1].Color);
            Assert.True(headerRichText.Runs[1].Bold);
            Assert.Equal(OfficeColor.FromRgb(37, 99, 235), headerRichText.Runs[2].Color);
            Assert.True(headerRichText.Runs[2].Italic);
            Assert.True(headerRichText.Runs[2].Underline);
            Assert.Equal(OfficeColor.FromRgb(22, 163, 74), footerRichText.Runs[1].Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Head", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("blue", svgText, StringComparison.Ordinal);
            Assert.Contains("Foot", svgText, StringComparison.Ordinal);
            Assert.Contains("Green", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterListMarkersThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;
            WordList headerList = document.HeaderFirstOrCreate.AddList(WordListStyle.Bulleted);
            headerList.AddItem("Header bullet");
            WordList footerList = document.FooterFirstOrCreate.AddList(WordListStyle.Numbered);
            footerList.AddItem("Footer number");
            document.AddParagraph("Body with header footer lists");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText headerText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Header bullet");
            OfficeDrawingText headerMarker = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Y == headerText.Y && text.X < headerText.X);
            Assert.False(string.IsNullOrWhiteSpace(headerMarker.Text));
            Assert.Equal("Symbol", headerMarker.Font.FamilyName);

            OfficeDrawingText footerText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Footer number");
            OfficeDrawingText footerMarker = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Y == footerText.Y && text.X < footerText.X);
            Assert.Equal("1.", footerMarker.Text);

            OfficeDrawingText body = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Body with header footer lists");
            Assert.True(headerText.Y < body.Y);
            Assert.True(footerText.Y > body.Y);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains(headerMarker.Text, svgText, StringComparison.Ordinal);
            Assert.Contains("Header bullet", svgText, StringComparison.Ordinal);
            Assert.Contains("1.", svgText, StringComparison.Ordinal);
            Assert.Contains("Footer number", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;

            byte[] headerPng = CreateSolidPng(24, 18, OfficeColor.FromRgb(220, 38, 38));
            using var headerStream = new MemoryStream(headerPng);
            document.HeaderFirstOrCreate
                .AddParagraph()
                .AddImage(headerStream, "header-red.png", 24, 18, description: "Header red marker");

            byte[] footerPng = CreateSolidPng(20, 16, OfficeColor.FromRgb(22, 163, 74));
            using var footerStream = new MemoryStream(footerPng);
            document.FooterFirstOrCreate
                .AddParagraph()
                .AddImage(footerStream, "footer-green.png", 20, 16, description: "Footer green marker");

            document.AddParagraph("Body with header footer images");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            IReadOnlyList<OfficeDrawingImage> images = snapshot.Drawing.Images;
            OfficeDrawingImage headerImage = images.Single(image => image.AlternativeText == "Header red marker");
            OfficeDrawingImage footerImage = images.Single(image => image.AlternativeText == "Footer green marker");
            OfficeDrawingText body = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Body with header footer images");
            Assert.True(headerImage.Projection.Y < body.Y);
            Assert.True(footerImage.Projection.Y > body.Y);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(220, 38, 38)) > 20);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(22, 163, 74)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterFloatingImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;

            byte[] headerPng = CreateSolidPng(32, 24, OfficeColor.FromRgb(220, 38, 38));
            using var headerStream = new MemoryStream(headerPng);
            WordImage headerImage = document.HeaderFirstOrCreate
                .AddParagraph()
                .InsertImage(headerStream, "header-floating-red.png", 32, 24, WrapTextImage.Square, "Header floating red marker");
            headerImage.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            headerImage.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(72D) };
            headerImage.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            headerImage.verticalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(24D) };
            document.HeaderFirstOrCreate.AddParagraph("Header floats beside marker");

            byte[] footerPng = CreateSolidPng(28, 20, OfficeColor.FromRgb(22, 163, 74));
            using var footerStream = new MemoryStream(footerPng);
            WordImage footerImage = document.FooterFirstOrCreate
                .AddParagraph()
                .InsertImage(footerStream, "footer-floating-green.png", 28, 20, WrapTextImage.Square, "Footer floating green marker");
            footerImage.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            footerImage.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(72D) };
            footerImage.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            footerImage.verticalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(758D) };
            document.FooterFirstOrCreate.AddParagraph("Footer floats beside marker");

            document.AddParagraph("Body with floating header footer images");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage renderedHeaderImage = snapshot.Drawing.Images.Single(image => image.AlternativeText == "Header floating red marker");
            OfficeDrawingImage renderedFooterImage = snapshot.Drawing.Images.Single(image => image.AlternativeText == "Footer floating green marker");
            OfficeDrawingText headerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Header floats beside marker");
            OfficeDrawingText footerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Footer floats beside marker");
            OfficeDrawingText body = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Body with floating header footer images");

            Assert.True(renderedHeaderImage.Projection.Y < body.Y);
            Assert.True(renderedFooterImage.Projection.Y > body.Y);
            Assert.True(headerText.X > renderedHeaderImage.Projection.X + renderedHeaderImage.Projection.Width);
            Assert.True(footerText.X > renderedFooterImage.Projection.X + renderedFooterImage.Projection.Width);
            Assert.True(headerText.Y >= renderedHeaderImage.Projection.Y);
            Assert.True(headerText.Y < renderedHeaderImage.Projection.Y + renderedHeaderImage.Projection.Height);
            Assert.True(footerText.Y >= renderedFooterImage.Projection.Y);
            Assert.True(footerText.Y < renderedFooterImage.Projection.Y + renderedFooterImage.Projection.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(220, 38, 38)) > 20);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(22, 163, 74)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterFloatingShapesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;

            WordShape headerShape = document.HeaderFirstOrCreate
                .AddParagraph()
                .AddShapeDrawing(ShapeType.RightArrow, 54D, 24D, 72D, 24D);
            headerShape.FillColor = OfficeColor.FromRgb(248, 113, 113);
            headerShape.StrokeColor = OfficeColor.FromRgb(127, 29, 29);
            headerShape.StrokeWeight = 2D;
            document.HeaderFirstOrCreate.AddParagraph("Header shape wraps beside marker");

            WordShape footerShape = document.FooterFirstOrCreate
                .AddParagraph()
                .AddShapeDrawing(ShapeType.RightArrow, 50D, 22D, 72D, 758D);
            footerShape.FillColor = OfficeColor.FromRgb(34, 197, 94);
            footerShape.StrokeColor = OfficeColor.FromRgb(21, 128, 61);
            footerShape.StrokeWeight = 2D;
            document.FooterFirstOrCreate.AddParagraph("Footer shape wraps beside marker");

            document.AddParagraph("Body with floating header footer shapes");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingShape renderedHeaderShape = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(248, 113, 113));
            OfficeDrawingShape renderedFooterShape = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Single(shape => shape.Shape.FillColor == OfficeColor.FromRgb(34, 197, 94));
            OfficeDrawingText headerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Header shape wraps beside marker");
            OfficeDrawingText footerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Footer shape wraps beside marker");
            OfficeDrawingText body = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Body with floating header footer shapes");

            Assert.True(renderedHeaderShape.Y < body.Y);
            Assert.True(renderedFooterShape.Y > body.Y);
            Assert.True(headerText.X >= renderedHeaderShape.X + renderedHeaderShape.Shape.Width - 0.5D);
            Assert.True(footerText.X >= renderedFooterShape.X + renderedFooterShape.Shape.Width - 0.5D);
            Assert.True(headerText.Y >= renderedHeaderShape.Y);
            Assert.True(headerText.Y < renderedHeaderShape.Y + renderedHeaderShape.Shape.Height);
            Assert.True(footerText.Y >= renderedFooterShape.Y);
            Assert.True(footerText.Y < renderedFooterShape.Y + renderedFooterShape.Shape.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(248, 113, 113)) > 20);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(34, 197, 94)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#F87171", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#22C55E", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterFloatingTextBoxesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;

            WordTextBox headerTextBox = document.HeaderFirstOrCreate
                .AddParagraph()
                .AddTextBox("Header floating text box", WrapTextImage.Square);
            headerTextBox.Width = (long)Math.Round(120D * 12700D);
            headerTextBox.Height = (long)Math.Round(32D * 12700D);
            headerTextBox.HorizontalPositionRelativeFrom = DW.HorizontalRelativePositionValues.Page;
            headerTextBox.HorizontalPositionOffset = int.Parse(PointsToEmusText(72D), CultureInfo.InvariantCulture);
            headerTextBox.VerticalPositionRelativeFrom = DW.VerticalRelativePositionValues.Page;
            headerTextBox.VerticalPositionOffset = int.Parse(PointsToEmusText(24D), CultureInfo.InvariantCulture);
            document.HeaderFirstOrCreate.AddParagraph("Header text wraps after text box");

            WordTextBox footerTextBox = document.FooterFirstOrCreate
                .AddParagraph()
                .AddTextBox("Footer floating text box", WrapTextImage.Square);
            footerTextBox.Width = (long)Math.Round(116D * 12700D);
            footerTextBox.Height = (long)Math.Round(32D * 12700D);
            footerTextBox.HorizontalPositionRelativeFrom = DW.HorizontalRelativePositionValues.Page;
            footerTextBox.HorizontalPositionOffset = int.Parse(PointsToEmusText(72D), CultureInfo.InvariantCulture);
            footerTextBox.VerticalPositionRelativeFrom = DW.VerticalRelativePositionValues.Page;
            footerTextBox.VerticalPositionOffset = int.Parse(PointsToEmusText(750D), CultureInfo.InvariantCulture);
            document.FooterFirstOrCreate.AddParagraph("Footer text wraps after text box");

            document.AddParagraph("Body with floating header footer text boxes");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText headerBoxText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text.Contains("Header floating", StringComparison.Ordinal));
            OfficeDrawingText footerBoxText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text.Contains("Footer floating", StringComparison.Ordinal));
            OfficeDrawingText headerAfterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Header text wraps after text box");
            OfficeDrawingText footerAfterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Footer text wraps after text box");
            OfficeDrawingText body = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Body with floating header footer text boxes");
            IReadOnlyList<OfficeDrawingShape> frames = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.StrokeColor == OfficeColor.Black)
                .ToList();
            OfficeDrawingShape headerFrame = frames.Single(shape => Math.Abs(shape.X - 72D) < 1D && shape.Y < body.Y);
            OfficeDrawingShape footerFrame = frames.Single(shape => Math.Abs(shape.X - 72D) < 1D && shape.Y > body.Y);

            Assert.True(headerBoxText.Y < body.Y);
            Assert.True(footerBoxText.Y > body.Y);
            Assert.True(headerAfterText.X >= headerFrame.X + headerFrame.Shape.Width - 0.5D);
            Assert.True(footerAfterText.X >= footerFrame.X + footerFrame.Shape.Width - 0.5D);
            Assert.True(headerAfterText.Y >= headerFrame.Y);
            Assert.True(headerAfterText.Y < headerFrame.Y + headerFrame.Shape.Height);
            Assert.True(footerAfterText.Y >= footerFrame.Y);
            Assert.True(footerAfterText.Y < footerFrame.Y + footerFrame.Shape.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(CountPixelsNear(rendered!, OfficeColor.Black) > 40);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Header", svgText, StringComparison.Ordinal);
            Assert.Contains("Footer", svgText, StringComparison.Ordinal);
            Assert.Contains("floating", svgText, StringComparison.Ordinal);
            Assert.Contains("box", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsHeaderFooterTablesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.DifferentFirstPage = true;

            WordTable headerTable = document.HeaderFirstOrCreate.AddTable(1, 1);
            headerTable.Rows[0].Cells[0].ShadingFillColor = OfficeColor.FromRgb(219, 234, 254);
            headerTable.Rows[0].Cells[0].AddParagraph("Header table cell", removeExistingParagraphs: true);

            WordTable footerTable = document.FooterFirstOrCreate.AddTable(1, 1);
            footerTable.Rows[0].Cells[0].ShadingFillColor = OfficeColor.FromRgb(220, 252, 231);
            footerTable.Rows[0].Cells[0].AddParagraph("Footer table cell", removeExistingParagraphs: true);

            document.AddParagraph("Body with header footer tables");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingText headerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Header table cell");
            OfficeDrawingText footerText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Footer table cell");
            OfficeDrawingText body = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == "Body with header footer tables");
            Assert.True(headerText.Y < body.Y);
            Assert.True(footerText.Y > body.Y);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape => shape.Shape.FillColor == OfficeColor.FromRgb(219, 234, 254));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape => shape.Shape.FillColor == OfficeColor.FromRgb(220, 252, 231));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(219, 234, 254)) > 20);
            Assert.True(CountPixelsNear(rendered!, OfficeColor.FromRgb(220, 252, 231)) > 20);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Header", svgText, StringComparison.Ordinal);
            Assert.Contains("table", svgText, StringComparison.Ordinal);
            Assert.Contains("Footer", svgText, StringComparison.Ordinal);
            Assert.Contains("cell", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordDocument_ReportsUnsupportedPageIndexes() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.AddTable(2, 2);

            OfficeImageExportResult secondPage = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { PageIndex = 1 });

            Assert.Contains(secondPage.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-page-index");
            Assert.DoesNotContain(secondPage.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-tables");
        }

        [Fact]
        public void WordDocument_ProjectsTightAndThroughWrappedImagesWithLimitedTextExclusionDiagnostics() {
            AssertLimitedSideWrappedImageProjection(WrapTextImage.Tight, "Tight");
            AssertLimitedSideWrappedImageProjection(WrapTextImage.Through, "Through");
        }

        [Fact]
        public void WordDocument_ScalesInlineImagesProportionallyWhenFittingContentWidth() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourcePng = CreateSolidPng(800, 400, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            document.AddParagraph().AddImage(imageStream, "wide.png", 800, 400, description: "Wide blue marker");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Wide blue marker", drawingImage.AlternativeText);
            Assert.True(drawingImage.Projection.Width < 800D);
            Assert.True(drawingImage.Projection.X + drawingImage.Projection.Width <= snapshot.Width);
            Assert.Equal(drawingImage.Projection.Width / 2D, drawingImage.Projection.Height, 1);
        }

        [Fact]
        public void WordDocument_ReservesProjectedInlineImageBoundsForPagination() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.PageSettings.Width = (UInt32Value)4000U;
            document.PageSettings.Height = (UInt32Value)3600U;
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Line one before rotated image.");
            document.AddParagraph("Line two before rotated image.");
            byte[] sourcePng = CreateSolidPng(240, 60, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage image = document.AddParagraph().InsertImage(imageStream, "rotated-inline-page.png", 240, 60, WrapTextImage.InLineWithText, "Rotated inline pagination marker");
            image.Rotation = 45;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Drawing.Images);
            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-pagination");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-image" && diagnostic.Source == "Rotated inline pagination marker");
        }

        [Fact]
        public void WordDocument_SkipsRotatedInlineImagesThatProjectOutsidePage() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            byte[] sourcePng = CreateSolidPng(420, 420, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage image = document.AddParagraph().InsertImage(imageStream, "rotated-inline.png", 420, 420, WrapTextImage.InLineWithText, "Rotated inline marker");
            image.Rotation = 45;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Drawing.Images);
            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-image" && diagnostic.Source == "Rotated inline marker");
        }

        private static void AssertLimitedSideWrappedImageProjection(WrapTextImage wrapText, string label) {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before " + label + " image");
            byte[] sourcePng = CreateSolidPng(48, 36, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            document.AddParagraph().AddImage(imageStream, label.ToLowerInvariant() + ".png", 48, 36, wrapText, label + " blue marker");
            document.AddParagraph("After " + label + " image wraps beside the marker.");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "limited-word-floating-image-wrap" && diagnostic.Source == label + " blue marker");
            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == "limited-word-floating-image-wrap" && diagnostic.Source == label + " blue marker");
            Assert.Contains(svg.Diagnostics, diagnostic => diagnostic.Code == "limited-word-floating-image-wrap" && diagnostic.Source == label + " blue marker");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-floating-image");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-body-element");
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal(label + " blue marker", drawingImage.AlternativeText);
            OfficeDrawingText afterText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "After " + label + " image wraps beside the marker.");
            Assert.True(afterText.X > drawingImage.Projection.X + drawingImage.Projection.Width);
            Assert.True(afterText.Y >= drawingImage.Projection.Y);
            Assert.True(afterText.Y < drawingImage.Projection.Y + drawingImage.Projection.Height);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsSquareWrappedAnchoredImagesThroughSharedDrawingTextExclusion() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before square image");
            byte[] sourcePng = CreateSolidPng(48, 36, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            document.AddParagraph().AddImage(imageStream, "square.png", 48, 36, WrapTextImage.Square, "Square blue marker");
            document.AddParagraph("After square image wraps beside the marker.");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Square blue marker", drawingImage.AlternativeText);
            Assert.Equal(36D, drawingImage.Projection.Width, 1);
            Assert.Equal(27D, drawingImage.Projection.Height, 1);
            OfficeDrawingText afterText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "After square image wraps beside the marker.");
            Assert.True(afterText.X > drawingImage.Projection.X + drawingImage.Projection.Width);
            Assert.True(afterText.Y >= drawingImage.Projection.Y);
            Assert.True(afterText.Y < drawingImage.Projection.Y + drawingImage.Projection.Height);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-floating-image");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsSquareWrappedImageSidePreferenceThroughTextExclusion() {
            AssertSquareWrappedImageSidePreference(DW.WrapTextValues.Left, "left", expectTextOnLeft: true);
            AssertSquareWrappedImageSidePreference(DW.WrapTextValues.Right, "right", expectTextOnLeft: false);
        }

        [Fact]
        public void WordDocument_ProjectsSquareWrappedImageDistanceMarginsThroughTextExclusion() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before square image distance margin");
            byte[] sourcePng = CreateSolidPng(48, 36, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "square-distance.png", 48, 36, WrapTextImage.Square, "square distance marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchored.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(220D) };
            anchored._Image.Anchor!.Elements<DW.WrapSquare>().Single().WrapText = DW.WrapTextValues.Right;
            anchored._Image.Anchor!.DistanceFromRight = new UInt32Value((uint)Math.Round(24D * 12700D));
            string afterTextValue = "After square image respects authored right distance.";
            document.AddParagraph(afterTextValue);

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            OfficeDrawingText afterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == afterTextValue);

            Assert.True(afterText.X >= drawingImage.Projection.X + drawingImage.Projection.Width + 23D);
            Assert.True(afterText.Y >= drawingImage.Projection.Y);
            Assert.True(afterText.Y < drawingImage.Projection.Y + drawingImage.Projection.Height);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        private static void AssertSquareWrappedImageSidePreference(DW.WrapTextValues wrapSide, string sideLabel, bool expectTextOnLeft) {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before " + sideLabel + " square image");
            byte[] sourcePng = CreateSolidPng(48, 36, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "square-side.png", 48, 36, WrapTextImage.Square, sideLabel + " square marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchored.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(220D) };
            anchored._Image.Anchor!.Elements<DW.WrapSquare>().Single().WrapText = wrapSide;
            string afterTextValue = "After " + sideLabel + " square image uses authored wrap side.";
            document.AddParagraph(afterTextValue);

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            OfficeDrawingText afterText = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Text == afterTextValue);

            if (expectTextOnLeft) {
                Assert.True(afterText.X < drawingImage.Projection.X);
            } else {
                Assert.True(afterText.X > drawingImage.Projection.X + drawingImage.Projection.Width);
            }

            Assert.True(afterText.Y >= drawingImage.Projection.Y);
            Assert.True(afterText.Y < drawingImage.Projection.Y + drawingImage.Projection.Height);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTopAndBottomAnchoredImagesThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before top-bottom image");
            byte[] sourcePng = CreateSolidPng(30, 20, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "topbottom.png", 30, 20, WrapTextImage.TopAndBottom, "Top bottom blue marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchored.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(144D) };
            anchored.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            anchored.verticalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(132D) };
            document.AddParagraph("After top-bottom image");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Top bottom blue marker", drawingImage.AlternativeText);
            Assert.Equal(144D, drawingImage.Projection.X, 1);
            Assert.Equal(132D, drawingImage.Projection.Y, 1);
            OfficeDrawingText afterText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "After top-bottom image");
            Assert.True(afterText.Y > drawingImage.Projection.Y + drawingImage.Projection.Height);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-floating-image");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTopAndBottomAnchoredImageBottomDistanceThroughFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before top-bottom image distance");
            byte[] sourcePng = CreateSolidPng(30, 20, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "topbottom-distance.png", 30, 20, WrapTextImage.TopAndBottom, "Top bottom distance marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchored.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(144D) };
            anchored.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            anchored.verticalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(132D) };
            anchored._Image.Anchor!.DistanceFromBottom = new UInt32Value((uint)Math.Round(24D * 12700D));
            string afterTextValue = "After top-bottom image respects bottom distance.";
            document.AddParagraph(afterTextValue);

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            OfficeDrawingText afterText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == afterTextValue);
            Assert.True(afterText.Y >= drawingImage.Projection.Y + drawingImage.Projection.Height + 23D);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-floating-image");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsNoWrapAnchoredImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            document.AddParagraph("Before anchored image");
            byte[] sourcePng = CreateSolidPng(32, 24, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "behind.png", 32, 24, WrapTextImage.BehindText, "Behind blue marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchored.horizontalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(126D) };
            anchored.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            anchored.verticalPosition.PositionOffset = new DW.PositionOffset { Text = PointsToEmusText(96D) };
            document.AddParagraph("After anchored image");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Behind blue marker", drawingImage.AlternativeText);
            Assert.Equal(126D, drawingImage.Projection.X, 1);
            Assert.Equal(96D, drawingImage.Projection.Y, 1);
            Assert.Equal(24D, drawingImage.Projection.Width, 1);
            Assert.Equal(18D, drawingImage.Projection.Height, 1);

            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Before anchored image");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "After anchored image");
            int imageIndex = snapshot.Drawing.Elements.ToList().IndexOf(drawingImage);
            int beforeTextIndex = snapshot.Drawing.Elements.ToList().FindIndex(element => element is OfficeDrawingText text && text.Text == "Before anchored image");
            int afterTextIndex = snapshot.Drawing.Elements.ToList().FindIndex(element => element is OfficeDrawingText text && text.Text == "After anchored image");
            Assert.InRange(imageIndex, 1, beforeTextIndex - 1);
            Assert.True(imageIndex < afterTextIndex);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-floating-image");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(37, 99, 235),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_AnchorsMarginRelativeFloatingImagesToPageMargin() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Normal;
            for (int i = 0; i < 5; i++) {
                document.AddParagraph("Body before anchor " + i.ToString(CultureInfo.InvariantCulture));
            }

            byte[] sourcePng = CreateSolidPng(32, 24, OfficeColor.FromRgb(37, 99, 235));
            using var imageStream = new MemoryStream(sourcePng);
            WordImage anchored = document.AddParagraph().InsertImage(imageStream, "margin-relative.png", 32, 24, WrapTextImage.Square, "Margin relative marker");
            anchored.horizontalPosition.RelativeFrom = DW.HorizontalRelativePositionValues.Margin;
            anchored.horizontalPosition.HorizontalAlignment = new DW.HorizontalAlignment { Text = "left" };
            anchored.verticalPosition.RelativeFrom = DW.VerticalRelativePositionValues.Margin;
            anchored.verticalPosition.VerticalAlignment = new DW.VerticalAlignment { Text = "top" };

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Margin relative marker", drawingImage.AlternativeText);
            Assert.Equal(72D, drawingImage.Projection.X, 1);
            Assert.Equal(72D, drawingImage.Projection.Y, 1);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            byte[] sourcePng = CreateSolidPng(18, 18, OfficeColor.FromRgb(220, 38, 38));
            using var imageStream = new MemoryStream(sourcePng);
            table.Rows[0].Cells[0].AddParagraph(removeExistingParagraphs: true).AddImage(imageStream, "cell.png", 18, 18, description: "Cell red marker");

            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-table-image");
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("Cell red marker", drawingImage.AlternativeText);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(
                OfficeColor.FromRgb(220, 38, 38),
                rendered!.GetPixel((int)(drawingImage.Projection.X + (drawingImage.Projection.Width / 2D)), (int)(drawingImage.Projection.Y + (drawingImage.Projection.Height / 2D))));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellTextBelowImagesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            byte[] sourcePng = CreateSolidPng(24, 24, OfficeColor.FromRgb(220, 38, 38));
            using var imageStream = new MemoryStream(sourcePng);
            cell.AddParagraph(removeExistingParagraphs: true).AddImage(imageStream, "cell-image.png", 24, 24, description: "Cell stacked image");
            cell.AddParagraph().AddText("Text below image");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();

            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            OfficeDrawingText text = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), item => item.Text == "Text below image");
            Assert.True(text.Y >= drawingImage.Projection.Y + drawingImage.Projection.Height, $"Expected table cell text below image, got text Y {text.Y} and image bottom {drawingImage.Projection.Y + drawingImage.Projection.Height}.");
        }

        [Fact]
        public void WordImageExportOptionsReuseSharedOfficeImageExportOptions() {
            WordImageExportOptions options = new WordImageExportOptions {
                Scale = 1.5D,
                BackgroundColor = OfficeColor.AliceBlue,
                IncludeDocumentContent = false
            };

            Assert.IsAssignableFrom<OfficeImageExportOptions>(options);
            Assert.Equal(1.5D, options.Scale);
            Assert.Equal(OfficeColor.AliceBlue, options.BackgroundColor);

            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { Scale = 0D }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { PageIndex = -1 }));
        }

        private static byte[] CreateSolidPng(int width, int height, OfficeColor color) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, color);
            return OfficePngWriter.Encode(image);
        }

        private static byte[] CreateSplitPng(int width, int height, OfficeColor leftColor, OfficeColor rightColor) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, leftColor);
            for (int y = 0; y < height; y++) {
                for (int x = width / 2; x < width; x++) {
                    image.SetPixel(x, y, rightColor);
                }
            }

            return OfficePngWriter.Encode(image);
        }

        private static byte[] CreateBmp24(int width, int height, IReadOnlyList<OfficeColor> pixels, bool topDown = false) {
            int rowStride = ((width * 24) + 31) / 32 * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, topDown ? -height : height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 24);

            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 3);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                }
            }

            return bytes;
        }

        private static byte[] CreateBmp32(int width, int height, IReadOnlyList<OfficeColor> pixels) {
            int rowStride = width * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 32);

            for (int y = 0; y < height; y++) {
                int sourceY = height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 4);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                    bytes[offset + 3] = color.A;
                }
            }

            return bytes;
        }

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

        private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static Run CreateWordTextBoxRun(string text, string hexColor, bool bold = false, bool italic = false, bool underline = false) {
            var properties = new RunProperties(
                new RunFonts { Ascii = "Aptos", HighAnsi = "Aptos" },
                new FontSize { Val = "22" },
                new Color { Val = hexColor });
            if (bold) {
                properties.Append(new Bold());
            }

            if (italic) {
                properties.Append(new Italic());
            }

            if (underline) {
                properties.Append(new Underline { Val = UnderlineValues.Single });
            }

            return new Run(properties, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        }

        private static string PointsToEmusText(double points) =>
            ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);

        private static int CountOccurrences(string text, string value) {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, StringComparison.OrdinalIgnoreCase)) >= 0) {
                count++;
                index += value.Length;
            }

            return count;
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (Math.Abs(actual.R - expected.R) <= 8 &&
                        Math.Abs(actual.G - expected.G) <= 8 &&
                        Math.Abs(actual.B - expected.B) <= 8 &&
                        Math.Abs(actual.A - expected.A) <= 8) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static void SetThemeColor(WordDocument document, string themeColor, string hexColor) {
            A.ColorScheme scheme = document.MainDocumentPartRoot.ThemePart!.Theme!.ThemeElements!.ColorScheme!;
            OpenXmlCompositeElement replacement = themeColor switch {
                "accent1" => new A.Accent1Color(new A.RgbColorModelHex { Val = hexColor }),
                "accent2" => new A.Accent2Color(new A.RgbColorModelHex { Val = hexColor }),
                "accent3" => new A.Accent3Color(new A.RgbColorModelHex { Val = hexColor }),
                _ => throw new ArgumentOutOfRangeException(nameof(themeColor), themeColor, "Unsupported test theme color.")
            };

            OpenXmlCompositeElement? existing = themeColor switch {
                "accent1" => scheme.GetFirstChild<A.Accent1Color>(),
                "accent2" => scheme.GetFirstChild<A.Accent2Color>(),
                "accent3" => scheme.GetFirstChild<A.Accent3Color>(),
                _ => null
            };
            existing?.InsertAfterSelf(replacement);
            existing?.Remove();
        }

    }
}
