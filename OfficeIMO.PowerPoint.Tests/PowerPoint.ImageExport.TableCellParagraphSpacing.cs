using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsTableCellParagraphSpacingThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 160, 90);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Cell spaced first";
            cell.FontSize = 12;
            cell.PaddingLeftPoints = 0D;
            cell.PaddingTopPoints = 0D;
            cell.PaddingRightPoints = 0D;
            cell.PaddingBottomPoints = 0D;
            PowerPointParagraph first = cell.Paragraphs.Single();
            first.SetSpaceAfterPoints(16);
            PowerPointParagraph second = cell.AddParagraph("Cell spaced second");
            second.SetSpaceBeforePoints(4);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            OfficeDrawingText firstText = SingleText(snapshot, "Cell spaced first");
            OfficeDrawingText secondText = SingleText(snapshot, "Cell spaced second");
            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), 19.9D, 20.1D);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.IndexOf('\n') >= 0);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell", svgText, StringComparison.Ordinal);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellLineSpacingThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 160, 90);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Cell line spacing";
            cell.FontSize = 12;
            cell.PaddingLeftPoints = 0D;
            cell.PaddingTopPoints = 0D;
            cell.PaddingRightPoints = 0D;
            cell.PaddingBottomPoints = 0D;
            PowerPointParagraph first = cell.Paragraphs.Single();
            first.SetLineSpacingPoints(22);
            cell.AddParagraph("Cell next line");

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(140, image!.Height);

            OfficeDrawingText firstText = SingleText(snapshot, "Cell line spacing");
            OfficeDrawingText secondText = SingleText(snapshot, "Cell next line");
            Assert.Equal(22D, firstText.LineHeight);
            Assert.InRange(firstText.Height, 21.9D, 22.1D);
            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), -0.1D, 0.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell", svgText, StringComparison.Ordinal);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsDefaultMultiParagraphTableCellThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 160, 90);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = 12;
            cell.PaddingLeftPoints = 0D;
            cell.PaddingTopPoints = 0D;
            cell.PaddingRightPoints = 0D;
            cell.PaddingBottomPoints = 0D;
            cell.SetParagraphs(new[] { "Cell first", "Cell second" });

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Equal(2, snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Count(text => text.Text.StartsWith("Cell ", StringComparison.Ordinal)));
            Assert.NotNull(SingleText(snapshot, "Cell first"));
            Assert.NotNull(SingleText(snapshot, "Cell second"));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell first", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell second", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("Cell firstCell second", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsScaledGroupedTableCellTextMetricsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 80, 30);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Grouped";
            cell.FontSize = 10;
            cell.PaddingLeftPoints = 3;
            cell.PaddingTopPoints = 2;
            cell.PaddingRightPoints = 4;
            cell.PaddingBottomPoints = 1;

            PowerPointAutoShape anchor = slide.AddRectanglePoints(110, 20, 10, 10);
            slide.GroupShapes(new PowerPointShape[] { table, anchor }, "Scaled table group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(200);
            transform.Extents.Cy = PowerPointUnits.FromPoints(60);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(100);
            transform.ChildExtents.Cy = PowerPointUnits.FromPoints(30);
            slide.SlidePart.Slide.Save();

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingText text = SingleText(snapshot, "Grouped");
            Assert.Equal(20D, text.Font.Size, 6);
            Assert.Equal(6D, text.Padding.Left, 6);
            Assert.Equal(4D, text.Padding.Top, 6);
            Assert.Equal(8D, text.Padding.Right, 6);
            Assert.Equal(2D, text.Padding.Bottom, 6);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellListMarkersThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(260, 180);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 200, 120);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = 12;
            cell.PaddingLeftPoints = 0D;
            cell.PaddingTopPoints = 0D;
            cell.PaddingRightPoints = 0D;
            cell.PaddingBottomPoints = 0D;
            cell.SetBullets(new[] { "Cell bullet", "Cell second" });
            cell.AddNumberedList(new[] { "Cell number", "Cell next" }, startAt: 5);
            cell.AddBullets(new[] { "Cell nested" }, level: 1);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(260, image!.Width);

            OfficeDrawingRichText[] richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(5, richTexts.Length);
            Assert.Contains(richTexts[0].Runs, run => run.Text == "\u2022 ");
            Assert.Contains(richTexts[2].Runs, run => run.Text == "5. ");
            Assert.Contains(richTexts[3].Runs, run => run.Text == "6. ");
            Assert.Contains(richTexts[4].Runs, run => run.Text == "\u2022 ");
            Assert.Contains("Cell bullet", richTexts[0].PlainText, StringComparison.Ordinal);
            Assert.Contains("Cell next", richTexts[3].PlainText, StringComparison.Ordinal);
            Assert.Contains("Cell nested", richTexts[4].PlainText, StringComparison.Ordinal);
            Assert.Equal(0D, richTexts[0].ParagraphIndent.FirstLineOffset);
            Assert.Equal(18D, richTexts[0].ParagraphIndent.ContinuationLineOffset);
            Assert.Equal(18D, richTexts[4].ParagraphIndent.FirstLineOffset);
            Assert.Equal(36D, richTexts[4].ParagraphIndent.ContinuationLineOffset);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell bullet", svgText, StringComparison.Ordinal);
            Assert.Contains("5.", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell next", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell nested", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsSingleTableCellBulletThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 160, 70);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = 12;
            cell.PaddingLeftPoints = 0D;
            cell.PaddingTopPoints = 0D;
            cell.PaddingRightPoints = 0D;
            cell.PaddingBottomPoints = 0D;
            cell.SetBullets(new[] { "Only bullet" });

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Contains(richText.Runs, run => run.Text == "\u2022 ");
            Assert.Contains("Only bullet", richText.PlainText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Only bullet", svgText, StringComparison.Ordinal);
        }
    }
}
