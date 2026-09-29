using System;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class WordToMarkdownConverter {
        private bool TryCreateChartSvgFallbackBlock(
            WordChart chart,
            WordToMarkdownOptions options,
            out IMarkdownBlock block) {
            block = null!;

            if (!chart.TryGetOfficeSnapshot(out var snapshot)) {
                options.OnWarning?.Invoke("Word chart could not be rendered as an SVG Markdown image because its cached chart data could not be read.");
                return false;
            }

            try {
                OfficeChartSnapshot officeSnapshot = snapshot;
                OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(officeSnapshot, useMinimumCanvas: false);
                if (rendering.QualityReport.HasIssues) {
                    options.OnWarning?.Invoke("Rendered Word chart '" + GetChartDisplayName(snapshot) + "' with shared drawing quality warnings: " + FormatQualityIssues(rendering.QualityReport));
                }

                byte[] svgBytes = OfficeDrawingSvgExporter.ToSvgBytes(rendering.Drawing);
                string displayName = GetChartDisplayName(snapshot);
                string source = options.VisualFallbackMode == MarkdownVisualFallbackMode.SvgFile
                    ? WriteVisualFallbackSvgResource(svgBytes, displayName, options)
                    : "data:image/svg+xml;base64," + System.Convert.ToBase64String(svgBytes);
                string alt = string.IsNullOrWhiteSpace(snapshot.Title) ? "Word chart" : snapshot.Title!;
                var sequence = new InlineSequence { AutoSpacing = false };
                sequence.AddRaw(new ImageInline(alt, source, title: null, plainAlt: alt));
                block = new ParagraphBlock(sequence);
                options.OnWarning?.Invoke("Rendered Word chart '" + displayName + "' as an SVG Markdown image fallback.");
                return true;
            } catch (Exception ex) {
                options.OnWarning?.Invoke("Word chart could not be rendered as an SVG Markdown image fallback. " + ex.Message);
                return false;
            }
        }

    }
}
