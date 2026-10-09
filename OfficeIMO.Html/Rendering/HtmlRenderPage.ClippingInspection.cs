using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    // Inspect each scene clip in its own coordinate space, before ancestor effects. This
    // includes intentional image/background crops but excludes the automatic output clip.
    internal IReadOnlyList<HtmlDiagnostic> InspectClipping(int maximumWidth, int maximumHeight,
        CancellationToken cancellationToken) {
        var diagnostics = new List<HtmlDiagnostic>();
        int inspectedClips = 0;
        Visit(_scene);
        return diagnostics.AsReadOnly();

        void Visit(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (HtmlRenderVisual visual in visuals) {
                cancellationToken.ThrowIfCancellationRequested();
                if (visual is HtmlRenderClipGroup || visual is HtmlRenderPathClipGroup) {
                    if (++inspectedClips > 1024) throw new NotSupportedException("Clipping inspection exceeds its 1024-clip work limit.");
                }
                if (visual is HtmlRenderClipGroup clip) {
                    // Expand only the unrestricted axes. Descendant clips and transforms remain
                    // part of the measured child drawing; this clip itself is intentionally absent.
                    var bounds = ResolveDrawingBufferBounds(clip.Visuals, 0D, 0D, _fonts, cancellationToken);
                    double left = clip.ClipHorizontal ? clip.ClipX : bounds.Left;
                    double top = clip.ClipVertical ? clip.ClipY : bounds.Top;
                    double right = clip.ClipHorizontal ? clip.ClipX + clip.ClipWidth : bounds.Right;
                    double bottom = clip.ClipVertical ? clip.ClipY + clip.ClipHeight : bounds.Bottom;
                    OfficeDrawingQualityReport report = InspectBounds(clip.Visuals, left, top,
                        Math.Max(0.01D, right - left), Math.Max(0.01D, bottom - top), maximumWidth, maximumHeight, cancellationToken);
                    int clipped = report.Issues.Count(issue => issue.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds);
                    if (clipped > 0) diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Html", HtmlRenderDiagnosticCodes.ClippedElementBounds,
                        "Rendered element bounds extend beyond a rectangular scene clip. The crop may be intentional.",
                        HtmlDiagnosticSeverity.Info, clip.Source,
                        string.Format(CultureInfo.InvariantCulture, "x={0};y={1};width={2};height={3};clipHorizontal={4};clipVertical={5};outsideBoundsFindings={6}",
                            clip.ClipX, clip.ClipY, clip.ClipWidth, clip.ClipHeight, clip.ClipHorizontal, clip.ClipVertical, clipped)));
                } else if (visual is HtmlRenderPathClipGroup path) {
                    diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Html", HtmlRenderDiagnosticCodes.ClipGeometryNotInspected,
                        "Path-shaped clipping requires separate geometry inspection; rectangular bounds cannot establish whether it hides content.",
                        HtmlDiagnosticSeverity.Warning, path.Source));
                }
                Visit(InspectionChildren(visual));
            }
        }
    }
}
