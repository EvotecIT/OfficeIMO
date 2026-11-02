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
            using var presentation = PowerPointDocument.Create();
            presentation.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(180));
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointParagraph paragraph;
            if (tableCell) {
                var cell = slide.AddTable(1, 1, OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(280), OfficeLength.Points(130)).GetCell(0, 0);
                cell.FontSize = OfficeLength.Points(20);
                cell.Text = "Leading first";
                cell.PaddingLeft = cell.PaddingRight = cell.PaddingTop = cell.PaddingBottom = OfficeLength.Zero;
                paragraph = cell.Paragraphs[0];
            } else {
                var box = slide.AddTextBox("Leading first", OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(280), OfficeLength.Points(130));
                box.FontSize = OfficeLength.Points(20);
                box.TextAutoFit = PowerPointTextAutoFit.None;
                box.SetTextMargins(OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0));
                slide.SlidePart.Slide.Descendants<A.BodyProperties>().Single().Wrap =
                    wrap ? A.TextWrappingValues.Square : A.TextWrappingValues.None;
                paragraph = box.Paragraphs[0];
            }
            paragraph.Runs[0].FontSize = OfficeLength.Points(15);
            paragraph.AddLineBreak().AddText("Leading middle", run => run.FontSize = OfficeLength.Points(middleSize));
            paragraph.AddLineBreak().AddText("Leading last", run => run.FontSize = OfficeLength.Points(15));
            paragraph.SetLineSpacing(OfficeLength.Points(20.25));

            using var saved = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            PowerPointSlideVisualInfo snapshot = saved.Slides[0].CreateVisualSnapshot();
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.NotNull(SingleText(snapshot, "Leading middle"));
            double advance = SingleText(snapshot, "Leading first").Y - SingleText(snapshot, "Leading last").Y;
            Assert.InRange(advance, 40.4, 40.6);
        }

        [Fact]
        public void HardBreakLeadingReportsTextBeyondTheAuthoredFrame() {
            using var presentation = PowerPointDocument.Create();
            presentation.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(180));
            var box = presentation.AddSlide().AddTextBox("Heading", OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(280), OfficeLength.Points(80));
            box.FontSize = OfficeLength.Points(20);
            box.TextAutoFit = PowerPointTextAutoFit.None;
            box.SetTextMargins(OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0));
            var paragraph = box.AddParagraph("Leading first");
            paragraph.Runs[0].FontSize = OfficeLength.Points(15);
            paragraph.AddLineBreak().AddText("Leading middle", run => run.FontSize = OfficeLength.Points(15));
            paragraph.AddLineBreak().AddText("Leading last", run => run.FontSize = OfficeLength.Points(15));
            paragraph.SetLineSpacing(OfficeLength.Points(20.25));

            PowerPointSlideVisualInfo snapshot = presentation.Slides[0].CreateVisualSnapshot();
            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
            Assert.NotNull(SingleText(snapshot, "Leading last"));
        }

        [Theory]
        [InlineData(11, 14.25, false)]
        [InlineData(24, 10, true)]
        public void OverflowUsesPaintedTextRatherThanUnusedTrailingLeading(int fontSize, double leading, bool overflow) {
            using var presentation = PowerPointDocument.Create();
            presentation.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(180));
            var box = presentation.AddSlide().AddTextBox("01", OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(100), OfficeLength.Points(14));
            box.FontSize = OfficeLength.Points(fontSize);
            box.TextAutoFit = PowerPointTextAutoFit.None;
            box.SetTextMargins(OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0));
            box.Paragraphs[0].SetLineSpacing(OfficeLength.Points(leading));
            using var saved = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
            var snapshot = saved.Slides[0].CreateVisualSnapshot();
            Assert.Equal(overflow, snapshot.Diagnostics.Any(d => d.Code == "powerpoint-text-overflow"));
            Assert.NotNull(SingleText(snapshot, "01"));
        }
    }
}
