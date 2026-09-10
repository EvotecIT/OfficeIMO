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
        public void PowerPointSlide_ProjectsTextBoxParagraphSpacingThroughSharedDrawingFlow() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("PowerPoint spaced first", 20, 20, 190, 80);
            textBox.FontSize = 12;
            textBox.SetTextMarginsPoints(0, 0, 0, 0);
            textBox.Paragraphs[0].SetSpaceAfterPoints(18);
            PowerPointParagraph second = textBox.AddParagraph("PowerPoint spaced second");
            second.SetSpaceBeforePoints(6);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            OfficeDrawingText firstText = SingleText(snapshot, "PowerPoint spaced first");
            OfficeDrawingText secondText = SingleText(snapshot, "PowerPoint spaced second");
            Assert.InRange(secondText.Y - firstText.Y, 38.9D, 39.1D);
            Assert.Equal(100D, firstText.Y + firstText.Height);
            Assert.Equal(100D, secondText.Y + secondText.Height);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.IndexOf('\n') >= 0);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("PowerPoint", svgText, StringComparison.Ordinal);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTextBoxLineSpacingThroughSharedDrawingFlow() {
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

            OfficeDrawingText firstText = SingleText(snapshot, "PowerPoint line spacing");
            OfficeDrawingText secondText = SingleText(snapshot, "PowerPoint next line");
            Assert.Equal(24D, firstText.LineHeight);
            Assert.InRange(secondText.Y - firstText.Y, 23.9D, 24.1D);
            Assert.Equal(100D, firstText.Y + firstText.Height);
            Assert.Equal(100D, secondText.Y + secondText.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("PowerPoint", svgText, StringComparison.Ordinal);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
        }

        private static OfficeDrawingText SingleText(PowerPointSlideVisualSnapshot snapshot, string text) =>
            snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(element => element.Text == text);
    }
}
