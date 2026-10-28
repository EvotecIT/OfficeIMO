using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Theory]
        [InlineData(false, "Direct")]
        [InlineData(true, "Direct")]
        [InlineData(false, "Theme")]
        [InlineData(true, "Theme")]
        [InlineData(false, "ListStyle")]
        [InlineData(true, "ListStyle")]
        [InlineData(false, "FollowText")]
        [InlineData(true, "FollowText")]
        [InlineData(false, "Reset")]
        [InlineData(true, "Reset")]
        public void SavedBulletColorUsesParagraphInheritanceWithoutRecoloringItsText(bool tableCell, string colorSource) {
            using var presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(320, 180);
            var slide = presentation.AddSlide();
            PowerPointParagraph paragraph;
            A.ListStyle listStyle;
            if (tableCell) {
                var cell = slide.AddTablePoints(1, 1, 20, 20, 280, 130).GetCell(0, 0);
                cell.Text = "Bullet body";
                cell.FontSize = 18;
                cell.Color = "111111";
                paragraph = cell.Paragraphs[0];
                listStyle = cell.Cell.TextBody!.ListStyle!;
            } else {
                var box = slide.AddTextBoxPoints("Bullet body", 20, 20, 280, 130);
                box.FontSize = 18;
                box.Color = "111111";
                paragraph = box.Paragraphs[0];
                listStyle = box.TextBody!.ListStyle!;
            }
            paragraph.SetBullet();
            var green = new A.BulletColor(new A.RgbColorModelHex { Val = "1DA155" });
            if (colorSource == "Theme") {
                var theme = slide.SlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme;
                theme.ThemeElements!.ColorScheme!.Accent1Color = new A.Accent1Color(new A.RgbColorModelHex { Val = "1DA155" });
                theme.Save();
                green = new A.BulletColor(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            }
            if (colorSource is "Direct" or "Theme") paragraph.Paragraph.ParagraphProperties!.PrependChild(green);
            if (colorSource is "ListStyle" or "Reset") listStyle.Append(new A.Level1ParagraphProperties(green));
            if (colorSource is "FollowText" or "Reset")
                paragraph.Paragraph.ParagraphProperties!.PrependChild(new A.BulletColorText());

            using var saved = PowerPointPresentation.Load(new MemoryStream(presentation.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            var snapshot = saved.Slides[0].CreateVisualSnapshot();
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeColor textColor = OfficeColor.FromRgb(17, 17, 17);
            var bullet = Assert.Single(snapshot.TextSpans, span => span.Text.Contains("\u2022"));
            Assert.Equal(colorSource is "Direct" or "Theme" or "ListStyle" ? OfficeColor.FromRgb(29, 161, 85) : textColor, bullet.Color);
            Assert.Equal(textColor, SingleText(snapshot, "Bullet body").Color);
        }
    }
}
