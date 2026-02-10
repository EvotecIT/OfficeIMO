using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests {
    public class Markdown_Renderer_FenceConversion_Tests {
        [Fact]
        public void MermaidEnabled_NoMermaidFence_DoesNotInjectMermaidNodes() {
            var options = new MarkdownRendererOptions();
            options.Mermaid.Enabled = true;
            var html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("# Title\n\nJust text.", options);

            Assert.DoesNotContain("class=\"mermaid\"", html);
        }

        [Fact]
        public void ChartEnabled_NoChartFence_DoesNotInjectCanvasNodes() {
            var options = new MarkdownRendererOptions();
            options.Chart.Enabled = true;
            var html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("# Title\n\nJust text.", options);

            Assert.DoesNotContain("class=\"omd-chart\"", html);
        }

        [Fact]
        public void MathEnabled_NoMathFence_DoesNotInjectMathWrapperNodes() {
            var options = new MarkdownRendererOptions();
            options.Math.Enabled = true;
            options.Math.EnableFencedMathBlocks = true;
            var html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("# Title\n\nJust text.", options);

            Assert.DoesNotContain("class=\"omd-math\"", html);
        }
    }
}
