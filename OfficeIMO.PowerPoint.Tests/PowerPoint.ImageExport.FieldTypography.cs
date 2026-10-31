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
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(340),OfficeIMO.OfficeLength.Points(150));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointParagraph textBoxParagraph = slide.AddTextBox("seed",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(135),global::OfficeIMO.OfficeLength.Points(60)).Paragraphs[0];
            textBoxParagraph.Paragraph.RemoveAllChildren<A.Run>();
            textBoxParagraph.AddField("text field", "custom", null, run => {
                run.Bold = true;
                run.Italic = true;
                run.FontName = "Arial";
                run.FontSize = global::OfficeIMO.OfficeLength.Points(15);
                run.UnderlineStyle = PowerPointUnderlineStyle.Double;
                run.Capitalization = PowerPointCapitalization.AllCaps;
            });

            PowerPointTableCell cell = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(175),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(140),global::OfficeIMO.OfficeLength.Points(55)).GetCell(0, 0);
            cell.Text = "seed";
            PowerPointParagraph cellParagraph = cell.Paragraphs[0];
            cellParagraph.Paragraph.RemoveAllChildren<A.Run>();
            cellParagraph.AddField("table field", "custom", null, run => {
                run.Bold = true;
                run.FontName = "Consolas";
                run.FontSize = global::OfficeIMO.OfficeLength.Points(13);
                run.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
                run.Capitalization = PowerPointCapitalization.SmallCaps;
            });

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            Assert.Contains("TEXT FIELD", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.Contains("TABLE FIELD", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            var textBoxSpans = snapshot.TextSpans.Where(span => span.X < 160 && !string.IsNullOrWhiteSpace(span.Text)).ToArray();
            var tableSpans = snapshot.TextSpans.Where(span => span.X >= 175 && !string.IsNullOrWhiteSpace(span.Text)).ToArray();
            Assert.NotEmpty(textBoxSpans);
            Assert.NotEmpty(tableSpans);
            Assert.All(textBoxSpans, span => {
                Assert.True(span.IsBold);
                Assert.True(span.IsItalic);
                Assert.Equal(15D, span.FontSize, 3);
            });
            Assert.All(tableSpans, span => {
                Assert.True(span.IsBold);
                Assert.Equal(13D, span.FontSize, 3);
            });
            Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
        }
        [Fact]
        public void FieldInlineRunMutationsPersistToThePresentationPackage() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
                    PowerPointParagraph paragraph = presentation.AddSlide().AddTextBox("Before").Paragraphs.Single();
                    paragraph.AddField("Field", "custom");
                    PowerPointParagraphInline authoredField = Assert.Single(paragraph.InlineNodes,
                        node => node.Kind == PowerPointParagraphInlineKind.Field);
                    authoredField.Run!.Text = "Changed";
                    authoredField.Run.Bold = true;
                    authoredField.Run.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
                    Assert.Equal("Changed", authoredField.Text);
                    presentation.Save();
                }

                using PowerPointDocument reloaded = PowerPointDocument.Load(path);
                PowerPointParagraphInline reloadedField = Assert.Single(reloaded.Slides.Single().TextBoxes.Single()
                    .Paragraphs.Single().InlineNodes, node => node.Kind == PowerPointParagraphInlineKind.Field);
                Assert.Equal("Changed", reloadedField.Text);
                Assert.True(reloadedField.Run!.Bold);
                Assert.Equal(PowerPointUnderlineStyle.Wavy, reloadedField.Run.UnderlineStyle);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void PowerPointSlide_ImageExportMaterializesNativeCapsForTextBoxesAndTableCells() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(140));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointParagraph textBoxParagraph = slide.AddTextBox("izmir",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(130),global::OfficeIMO.OfficeLength.Points(60)).Paragraphs[0];
            PowerPointTextRun textBoxRun = textBoxParagraph.Runs[0];
            textBoxRun.Language = "tr-TR";
            textBoxRun.Capitalization = PowerPointCapitalization.AllCaps;
            textBoxParagraph.AddLineBreak().AddField("TextBoxField", "custom");

            PowerPointTableCell cell = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(170),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(50)).GetCell(0, 0);
            cell.Text = "small caps";
            cell.Runs[0].Capitalization = PowerPointCapitalization.SmallCaps;
            cell.Paragraphs[0].AddLineBreak().AddField("CellField", "custom");

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            string text = GetSnapshotPlainText(slide.CreateVisualSnapshot());

            Assert.Contains("TextBoxField", text, StringComparison.Ordinal);
            Assert.Contains("İZMİR", text, StringComparison.Ordinal);
            Assert.Contains("SMALL CAPS", text, StringComparison.Ordinal);
            Assert.Contains("CellField", text, StringComparison.Ordinal);
            Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
        }

        [Fact]
        public void PowerPointSlide_ImageExportResolvesParagraphCapsAndReportsNoncanonicalBaselines() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(150));
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointParagraph textParagraph = slide.AddTextBox("izmir",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(140),global::OfficeIMO.OfficeLength.Points(50)).Paragraphs[0];
            textParagraph.Paragraph.ParagraphProperties = new A.ParagraphProperties(
                new A.DefaultRunProperties { Capital = A.TextCapsValues.All, Language = "tr-TR" });
            textParagraph.Runs[0].BaselinePercent = 5D;

            PowerPointTableCell cell = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(180),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(140),global::OfficeIMO.OfficeLength.Points(50)).GetCell(0, 0);
            cell.Text = "small caps";
            PowerPointParagraph cellParagraph = cell.Paragraphs[0];
            cellParagraph.Paragraph.ParagraphProperties = new A.ParagraphProperties(
                new A.DefaultRunProperties { Capital = A.TextCapsValues.Small, Language = "en-US" });
            cellParagraph.Runs[0].BaselinePercent = 80D;

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            string text = GetSnapshotPlainText(slide.CreateVisualSnapshot());

            Assert.Contains("İZMİR", text, StringComparison.Ordinal);
            Assert.Contains("SMALL CAPS", text, StringComparison.Ordinal);
            Assert.Equal(2, svg.Diagnostics.Count(diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated));
            Assert.Equal(2, png.Diagnostics.Count(diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated));
            Assert.Single(svg.Diagnostics, diagnostic => diagnostic.Code == PowerPointImageExportDiagnosticCodes.SmallCapsApproximated);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
        }

    }
}
