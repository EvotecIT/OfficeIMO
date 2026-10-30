using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsTableCellParagraphSpacingInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(220),OfficeIMO.OfficeLength.Points(140));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(90));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Cell spaced first";
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(0D);
            PowerPointParagraph first = cell.Paragraphs.Single();
            PowerPointParagraph second = cell.AddParagraph("Cell spaced second");
            PowerPointSlideVisualInfo unspaced = slide.CreateVisualSnapshot();
            first.SetSpaceAfter(global::OfficeIMO.OfficeLength.Points(16));
            second.SetSpaceBefore(global::OfficeIMO.OfficeLength.Points(4));

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            PowerPointTextSpanInfo firstText = SingleText(snapshot, "Cell spaced first");
            PowerPointTextSpanInfo secondText = SingleText(snapshot, "Cell spaced second");
            double baselineGap = SingleText(unspaced, "Cell spaced first").Y - SingleText(unspaced, "Cell spaced second").Y;
            Assert.InRange(firstText.Y - secondText.Y - baselineGap, 19.9D, 20.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellLineSpacingInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(220),OfficeIMO.OfficeLength.Points(140));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(90));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Cell line spacing";
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(0D);
            PowerPointParagraph first = cell.Paragraphs.Single();
            first.SetLineSpacing(global::OfficeIMO.OfficeLength.Points(22));
            cell.AddParagraph("Cell next line");

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(140, image!.Height);

            PowerPointTextSpanInfo firstText = SingleText(snapshot, "Cell line spacing");
            PowerPointTextSpanInfo secondText = SingleText(snapshot, "Cell next line");
            Assert.InRange(firstText.Y - secondText.Y, 21.9D, 22.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsDefaultMultiParagraphTableCellInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(220),OfficeIMO.OfficeLength.Points(140));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(90));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(0D);
            cell.SetParagraphs(new[] { "Cell first", "Cell second" });

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.NotNull(SingleText(snapshot, "Cell first"));
            Assert.NotNull(SingleText(snapshot, "Cell second"));

            Assert.Contains("Cell first", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("Cell second", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.True(SingleText(snapshot, "Cell first").Y > SingleText(snapshot, "Cell second").Y);
        }

        [Fact]
        public void PowerPointSlide_ProjectsScaledGroupedTableCellTextMetricsInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(240),OfficeIMO.OfficeLength.Points(140));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(80),global::OfficeIMO.OfficeLength.Points(30));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Grouped";
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(10);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(3);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(2);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(4);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(1);

            PowerPointAutoShape anchor = slide.AddRectangle(OfficeIMO.OfficeLength.Points(110), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10));
            slide.GroupShapes(new PowerPointShape[] { table, anchor }, "Scaled table group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = global::OfficeIMO.OfficeLength.Points(200).ToEmus();
            transform.Extents.Cy = global::OfficeIMO.OfficeLength.Points(60).ToEmus();
            transform.ChildExtents!.Cx = global::OfficeIMO.OfficeLength.Points(100).ToEmus();
            transform.ChildExtents.Cy = global::OfficeIMO.OfficeLength.Points(30).ToEmus();
            slide.SlidePart.Slide.Save();

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            PowerPointTextSpanInfo text = SingleText(snapshot, "Grouped");
            // Scene text geometry includes the group's vertical scale.
            Assert.Equal(20D, text.FontSize, 3);
            Assert.InRange(text.X, 25.9D, 26.1D);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellListMarkersInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(260),OfficeIMO.OfficeLength.Points(180));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(120));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(0D);
            cell.SetBullets(new[] { "Cell bullet", "Cell second" });
            cell.AddNumberedList(new[] { "Cell number", "Cell next" }, startAt: 5);
            cell.AddBullets(new[] { "Cell nested" }, level: 1);

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics, "table-list-layout-approximation");
            AssertNoUnexpectedDiagnostics(png.Diagnostics, "table-list-layout-approximation");
            AssertNoUnexpectedDiagnostics(svg.Diagnostics, "table-list-layout-approximation");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(260, image!.Width);

            string text = GetSnapshotPlainText(snapshot);
            foreach (string expected in new[] { "Cell bullet", "Cell second", "5.", "6.", "Cell number", "Cell next", "Cell nested" })
                Assert.Contains(expected, text, StringComparison.Ordinal);
            Assert.Contains("\u2022", text, StringComparison.Ordinal);


        }

        [Fact]
        public void PowerPointSlide_ProjectsSingleTableCellBulletInNativeTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(220),OfficeIMO.OfficeLength.Points(120));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(70));
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(0D);
            cell.PaddingBottom = global::OfficeIMO.OfficeLength.Points(0D);
            cell.SetBullets(new[] { "Only bullet" });

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics, "table-list-layout-approximation");
            AssertNoUnexpectedDiagnostics(png.Diagnostics, "table-list-layout-approximation");
            AssertNoUnexpectedDiagnostics(svg.Diagnostics, "table-list-layout-approximation");
            Assert.Contains("\u2022", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("Only bullet", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);


        }
    }
}
