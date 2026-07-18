using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordImageExportTests {
        [Fact]
        public void WordDocument_ProjectsTableCellParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            ClearDocumentDefaultParagraphSpacing(document);
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Cell spacing first";
            cell.Paragraphs[0].LineSpacingAfterPoints = 18D;
            WordParagraph second = cell.AddParagraph("Cell spacing second");
            second.LineSpacingBeforePoints = 6D;
            second.LineSpacingAfterPoints = 0D;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText firstText = SingleText(snapshot, "Cell spacing first");
            OfficeDrawingText secondText = SingleText(snapshot, "Cell spacing second");
            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), 23.9D, 24.1D);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.IndexOf('\n') >= 0);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell spacing first", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell spacing second", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellContextualParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            ClearDocumentDefaultParagraphSpacing(document);
            const string contextualStyleId = "ImageTableCellContextualSpacing";
            const string otherStyleId = "ImageTableCellContextualSpacingOther";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Table Cell Contextual Spacing" },
                    new StyleParagraphProperties(
                        new SpacingBetweenLines { After = "480" },
                        new ContextualSpacing())) {
                    Type = StyleValues.Paragraph,
                    StyleId = contextualStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Image Table Cell Contextual Spacing Other" },
                    new StyleParagraphProperties(new SpacingBetweenLines { After = "0" })) {
                    Type = StyleValues.Paragraph,
                    StyleId = otherStyleId,
                    CustomStyle = true
                });
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Cell contextual first";
            cell.Paragraphs[0].SetStyleId(contextualStyleId);
            cell.AddParagraph("Cell contextual second").SetStyleId(contextualStyleId);
            cell.AddParagraph("Cell contextual third").SetStyleId(otherStyleId);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Height, image!.Height);

            OfficeDrawingText firstText = SingleText(snapshot, "Cell contextual first");
            OfficeDrawingText secondText = SingleText(snapshot, "Cell contextual second");
            OfficeDrawingText thirdText = SingleText(snapshot, "Cell contextual third");
            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), -0.1D, 0.1D);
            Assert.InRange(thirdText.Y - (secondText.Y + secondText.Height), 23.9D, 24.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell contextual first", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell contextual third", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsTableCellListMarkersThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            WordTable table = document.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 4200;
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.ColumnWidth = new List<int> { 4200 };
            WordTableCell cell = table.Rows[0].Cells[0];
            WordList bullets = cell.AddList(WordListStyle.Bulleted);
            bullets.AddItem("Cell bullet");
            WordList numbered = cell.AddList(WordListStyle.Numbered);
            numbered.AddItem("Cell number");

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText bulletText = SingleText(snapshot, "Cell bullet");
            OfficeDrawingText bulletMarker = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Y == bulletText.Y && text.X < bulletText.X);
            Assert.False(string.IsNullOrWhiteSpace(bulletMarker.Text));
            Assert.Equal("Symbol", bulletMarker.Font.FamilyName);

            OfficeDrawingText numberText = SingleText(snapshot, "Cell number");
            OfficeDrawingText numberMarker = snapshot.Drawing.Elements
                .OfType<OfficeDrawingText>()
                .Single(text => text.Y == numberText.Y && text.X < numberText.X);
            Assert.Equal("1.", numberMarker.Text);
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.Contains("Cell bullet\nCell number", StringComparison.Ordinal));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Cell bullet", svgText, StringComparison.Ordinal);
            Assert.Contains("Cell number", svgText, StringComparison.Ordinal);
            Assert.Contains("1.", svgText, StringComparison.Ordinal);
        }
    }
}
