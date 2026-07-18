using System;
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
        public void WordDocument_ProjectsDirectParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;

            WordParagraph before = document.AddParagraph("Before direct spacing");
            before.LineSpacingAfterPoints = 0D;
            WordParagraph spaced = document.AddParagraph("Direct spaced paragraph");
            spaced.LineSpacingBeforePoints = 18D;
            spaced.LineSpacingAfterPoints = 24D;
            WordParagraph after = document.AddParagraph("After direct spacing");
            after.LineSpacingAfterPoints = 0D;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText beforeText = SingleText(snapshot, "Before direct spacing");
            OfficeDrawingText spacedText = SingleText(snapshot, "Direct spaced paragraph");
            OfficeDrawingText afterText = SingleText(snapshot, "After direct spacing");

            Assert.InRange(spacedText.Y - (beforeText.Y + beforeText.Height), 17.9D, 18.1D);
            Assert.InRange(afterText.Y - (spacedText.Y + spacedText.Height), 23.9D, 24.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Direct spaced paragraph", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsStyleInheritedParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            const string styleId = "ImageParagraphSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Image Paragraph Spacing Style" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines { Before = "120", After = "360" })) {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph before = document.AddParagraph("Before styled spacing");
            before.LineSpacingAfterPoints = 0D;
            WordParagraph spaced = document.AddParagraph("Style spaced ");
            spaced.AddText("paragraph").SetBold();
            spaced.SetStyleId(styleId);
            WordParagraph after = document.AddParagraph("After styled spacing");
            after.LineSpacingAfterPoints = 0D;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Height, image!.Height);

            OfficeDrawingText beforeText = SingleText(snapshot, "Before styled spacing");
            OfficeDrawingRichText spacedText = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().Single(element => element.PlainText == "Style spaced paragraph");
            OfficeDrawingText afterText = SingleText(snapshot, "After styled spacing");

            Assert.InRange(spacedText.Y - (beforeText.Y + beforeText.Height), 5.9D, 6.1D);
            Assert.InRange(afterText.Y - (spacedText.Y + spacedText.Height), 17.9D, 18.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Style spaced", svgText, StringComparison.Ordinal);
            Assert.Contains("paragraph", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsDocumentDefaultParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            ParagraphPropertiesBaseStyle paragraphDefaults = document._wordprocessingDocument.MainDocumentPart!
                .StyleDefinitionsPart!
                .Styles!
                .DocDefaults!
                .GetFirstChild<ParagraphPropertiesDefault>()!
                .GetFirstChild<ParagraphPropertiesBaseStyle>()!;
            SpacingBetweenLines spacing = paragraphDefaults.GetFirstChild<SpacingBetweenLines>()!;
            spacing.Before = "120";
            spacing.After = "320";

            WordParagraph before = document.AddParagraph("Before defaults");
            before.LineSpacingAfterPoints = 0D;
            WordParagraph defaultSpaced = document.AddParagraph("Default spaced paragraph");
            WordParagraph overrideSpaced = document.AddParagraph("Override spaced paragraph");
            overrideSpaced.LineSpacingBeforePoints = 0D;
            overrideSpaced.LineSpacingAfterPoints = 0D;
            WordParagraph after = document.AddParagraph("After override");
            after.LineSpacingBeforePoints = 0D;
            after.LineSpacingAfterPoints = 0D;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText beforeText = SingleText(snapshot, "Before defaults");
            OfficeDrawingText defaultText = SingleText(snapshot, "Default spaced paragraph");
            OfficeDrawingText overrideText = SingleText(snapshot, "Override spaced paragraph");
            OfficeDrawingText afterText = SingleText(snapshot, "After override");

            Assert.InRange(defaultText.Y - (beforeText.Y + beforeText.Height), 5.9D, 6.1D);
            Assert.InRange(overrideText.Y - (defaultText.Y + defaultText.Height), 15.9D, 16.1D);
            Assert.InRange(afterText.Y - (overrideText.Y + overrideText.Height), -0.1D, 0.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Default spaced paragraph", svgText, StringComparison.Ordinal);
            Assert.Contains("Override spaced paragraph", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsDirectContextualParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            ClearDocumentDefaultParagraphSpacing(document);

            WordParagraph first = document.AddParagraph("Direct contextual first");
            first.LineSpacingAfterPoints = 24D;
            first._paragraph.ParagraphProperties!.Append(new ContextualSpacing());
            WordParagraph second = document.AddParagraph("Direct contextual second");
            second.LineSpacingBeforePoints = 12D;
            second.LineSpacingAfterPoints = 0D;

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);

            OfficeDrawingText firstText = SingleText(snapshot, "Direct contextual first");
            OfficeDrawingText secondText = SingleText(snapshot, "Direct contextual second");
            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), -0.1D, 0.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Direct contextual first", svgText, StringComparison.Ordinal);
            Assert.Contains("Direct contextual second", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void WordDocument_ProjectsStyleInheritedContextualParagraphSpacingThroughImageFlow() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            document.Margins.Type = WordMargin.Narrow;
            ClearDocumentDefaultParagraphSpacing(document);
            const string contextualStyleId = "ImageContextualSpacing";
            const string otherStyleId = "ImageContextualSpacingOther";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Image Contextual Spacing" },
                    new StyleParagraphProperties(
                        new SpacingBetweenLines { After = "480" },
                        new ContextualSpacing())) {
                    Type = StyleValues.Paragraph,
                    StyleId = contextualStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Image Contextual Spacing Other" },
                    new StyleParagraphProperties(new SpacingBetweenLines { After = "0" })) {
                    Type = StyleValues.Paragraph,
                    StyleId = otherStyleId,
                    CustomStyle = true
                });

            document.AddParagraph("Style contextual first").SetStyleId(contextualStyleId);
            document.AddParagraph("Style contextual second").SetStyleId(contextualStyleId);
            document.AddParagraph("Style contextual third").SetStyleId(otherStyleId);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, new WordImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, new WordImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Height, image!.Height);

            OfficeDrawingText firstText = SingleText(snapshot, "Style contextual first");
            OfficeDrawingText secondText = SingleText(snapshot, "Style contextual second");
            OfficeDrawingText thirdText = SingleText(snapshot, "Style contextual third");

            Assert.InRange(secondText.Y - (firstText.Y + firstText.Height), -0.1D, 0.1D);
            Assert.InRange(thirdText.Y - (secondText.Y + secondText.Height), 23.9D, 24.1D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Style contextual first", svgText, StringComparison.Ordinal);
            Assert.Contains("Style contextual third", svgText, StringComparison.Ordinal);
        }

        private static void ClearDocumentDefaultParagraphSpacing(WordDocument document) {
            SpacingBetweenLines spacing = document._wordprocessingDocument.MainDocumentPart!
                .StyleDefinitionsPart!
                .Styles!
                .DocDefaults!
                .GetFirstChild<ParagraphPropertiesDefault>()!
                .GetFirstChild<ParagraphPropertiesBaseStyle>()!
                .GetFirstChild<SpacingBetweenLines>()!;
            spacing.Before = "0";
            spacing.After = "0";
        }

        private static OfficeDrawingText SingleText(WordDocumentVisualSnapshot snapshot, string text) =>
            snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(element => element.Text == text);
    }
}
