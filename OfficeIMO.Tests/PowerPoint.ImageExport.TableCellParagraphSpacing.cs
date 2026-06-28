using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsTableCellParagraphSpacingThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.Slides[0];

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

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
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
            PowerPointSlide slide = presentation.Slides[0];

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

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
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
        public void PowerPointSlide_ProjectsTableCellListMarkersThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(260, 180);
            PowerPointSlide slide = presentation.Slides[0];

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

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
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
            PowerPointSlide slide = presentation.Slides[0];

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

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
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
