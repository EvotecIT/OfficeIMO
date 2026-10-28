using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Theory]
        [InlineData(false, true, 15)]
        [InlineData(false, true, 10)]
        [InlineData(false, false, 10)]
        [InlineData(true, true, 15)]
        [InlineData(true, true, 10)]
        public void SavedAbsoluteLeadingSurvivesHardBreaksAndMixedRunSizes(bool tableCell, bool wrap, int middleSize) {
            using var presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(320, 180);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointParagraph paragraph;
            if (tableCell) {
                var cell = slide.AddTablePoints(1, 1, 20, 20, 280, 130).GetCell(0, 0);
                cell.FontSize = 20;
                cell.Text = "Leading first";
                cell.PaddingLeftPoints = cell.PaddingRightPoints = cell.PaddingTopPoints = cell.PaddingBottomPoints = 0;
                paragraph = cell.Paragraphs[0];
            } else {
                var box = slide.AddTextBoxPoints("Leading first", 20, 20, 280, 130);
                box.FontSize = 20;
                box.TextAutoFit = PowerPointTextAutoFit.None;
                box.SetTextMarginsPoints(0, 0, 0, 0);
                slide.SlidePart.Slide.Descendants<A.BodyProperties>().Single().Wrap =
                    wrap ? A.TextWrappingValues.Square : A.TextWrappingValues.None;
                paragraph = box.Paragraphs[0];
            }
            paragraph.Runs[0].FontSize = 15;
            paragraph.AddLineBreak().AddText("Leading middle", run => run.FontSize = middleSize);
            paragraph.AddLineBreak().AddText("Leading last", run => run.FontSize = 15);
            paragraph.SetLineSpacingPoints(20.25);

            using var saved = PowerPointPresentation.Load(new MemoryStream(presentation.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            PowerPointSlideVisualInfo snapshot = saved.Slides[0].CreateVisualSnapshot();
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.NotNull(SingleText(snapshot, "Leading middle"));
            double advance = SingleText(snapshot, "Leading first").Y - SingleText(snapshot, "Leading last").Y;
            Assert.InRange(advance, 40.4, 40.6);
        }

        [Fact]
        public void HardBreakLeadingReportsTextBeyondTheAuthoredFrame() {
            using var presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(320, 180);
            var box = presentation.AddSlide().AddTextBoxPoints("Heading", 20, 20, 280, 80);
            box.FontSize = 20;
            box.TextAutoFit = PowerPointTextAutoFit.None;
            box.SetTextMarginsPoints(0, 0, 0, 0);
            var paragraph = box.AddParagraph("Leading first");
            paragraph.Runs[0].FontSize = 15;
            paragraph.AddLineBreak().AddText("Leading middle", run => run.FontSize = 15);
            paragraph.AddLineBreak().AddText("Leading last", run => run.FontSize = 15);
            paragraph.SetLineSpacingPoints(20.25);

            PowerPointSlideVisualInfo snapshot = presentation.Slides[0].CreateVisualSnapshot();
            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
            Assert.NotNull(SingleText(snapshot, "Leading last"));
        }
    }
}
