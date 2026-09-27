using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WordToMarkdown_VisualFallbackUsesSharedBubbleAndCombinationProjection(bool bubble) {
            using var document = WordDocument.Create();
            var data = bubble ? new OfficeChartData(new[] { "1", "2" }, new[] {
                OfficeChartSeries.CreateBubble("Measured", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 20d }, OfficeColor.Parse("#224466")) }) :
                new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Columns", new[] { 100d, 200d }),
                    new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, OfficeColor.Parse("#224466"), null, true,
                        renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) });
            document.AddChart(bubble ? OfficeChartKind.Bubble : OfficeChartKind.ColumnClustered, data, "Shared chart");
            string markdown = document.ToMarkdown(new WordToMarkdownOptions { VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri });
            const string prefix = "data:image/svg+xml;base64,";
            int start = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(start >= 0, markdown);
            start += prefix.Length;
            string svg = Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(start, markdown.IndexOf(')', start) - start)));
            Assert.Contains("Shared chart", svg);
            Assert.Contains("#224466", svg);
            Assert.Contains(bubble ? "Measured" : "Ratio", svg);
        }
    }
}
