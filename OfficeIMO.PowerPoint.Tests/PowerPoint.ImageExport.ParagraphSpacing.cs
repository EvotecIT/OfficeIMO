using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsTextBoxParagraphSpacingInPdfTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("PowerPoint spaced first", 20, 20, 190, 80);
            textBox.FontSize = 12;
            textBox.SetTextMarginsPoints(0, 0, 0, 0);
            PowerPointParagraph second = textBox.AddParagraph("PowerPoint spaced second");
            PowerPointSlideVisualSnapshot unspaced = slide.CreateVisualSnapshot();
            textBox.Paragraphs[0].SetSpaceAfterPoints(18);
            second.SetSpaceBeforePoints(6);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            PdfTextSpan firstText = SingleText(snapshot, "PowerPoint spaced first");
            PdfTextSpan secondText = SingleText(snapshot, "PowerPoint spaced second");
            double baselineGap = SingleText(unspaced, "PowerPoint spaced first").Y - SingleText(unspaced, "PowerPoint spaced second").Y;
            Assert.InRange(firstText.Y - secondText.Y - baselineGap, 23.9D, 24.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("PowerPoint", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTextBoxLineSpacingInPdfTextSpans() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("PowerPoint line spacing", 20, 20, 190, 80);
            textBox.FontSize = 12;
            textBox.SetTextMarginsPoints(0, 0, 0, 0);
            textBox.Paragraphs[0].SetLineSpacingPoints(24);
            textBox.AddParagraph("PowerPoint next line");

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(120, image!.Height);

            PdfTextSpan firstText = SingleText(snapshot, "PowerPoint line spacing");
            PdfTextSpan secondText = SingleText(snapshot, "PowerPoint next line");
            Assert.InRange(firstText.Y - secondText.Y, 23.9D, 24.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("PowerPoint", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
        }

        private static PdfTextSpan SingleText(PowerPointSlideVisualSnapshot snapshot, string text) {
            // PDF text can split a logical label across runs, font subsets, or wrapped lines.
            string combined = GetSnapshotPlainText(snapshot);
            Assert.Contains(text, combined, StringComparison.Ordinal);
            int start = combined.IndexOf(text, StringComparison.Ordinal);
            int offset = 0;
            foreach (var segment in GetSnapshotTextSegments(snapshot)) {
                if (offset + segment.Text.Length > start) return segment.Span;
                offset += segment.Text.Length;
            }
            throw new InvalidOperationException("The matched text must have a source span.");
        }
    }
}
