using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    internal IReadOnlyList<HtmlDiagnostic> InspectTextInk(double width, double height, CancellationToken token,
        IOfficeTextShapingProvider? shapingProvider = null, string? shapingLanguage = null, bool isRegion = false) {
        var diagnostics = new List<HtmlDiagnostic>();
        var measurementDiagnostics = new List<OfficeImageExportDiagnostic>();
        var measurement = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: _fonts,
            textShapingProvider: shapingProvider, textShapingLanguage: shapingLanguage,
            diagnosticSink: measurementDiagnostics, cancellationToken: token);
        int runs = 0;
        Visit(_scene.Where(v => v.Source != "render-surface"), OfficeTransform.Identity);
        foreach (var diagnostic in measurementDiagnostics) diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Html",
            HtmlRenderDiagnosticCodes.TextInkNotInspected, "Text measurement used a diagnosed shaping fallback: " + diagnostic.Message,
            HtmlDiagnosticSeverity.Warning, diagnostic.Source, diagnostic.Code));
        return diagnostics.AsReadOnly();

        void Visit(IEnumerable<HtmlRenderVisual> visuals, OfficeTransform transform) {
            foreach (var visual in visuals) {
                token.ThrowIfCancellationRequested();
                if (visual is HtmlRenderEffectGroup effect) {
                    if (effect.Opacity > 0D) Visit(effect.Visuals, effect.Transform.Then(transform));
                } else if (visual is HtmlRenderClipGroup || visual is HtmlRenderPathClipGroup) {
                    if (ContainsText(InspectionChildren(visual))) NotInspected(visual, "Text inside an authored clip requires separate clipped-ink inspection.");
                } else if (visual is HtmlRenderDrawing) {
                    NotInspected(visual, "Text within an embedded vector drawing is outside positioned XHTML text inspection.");
                } else if (visual is HtmlRenderText original && original.Text.Length > 0) {
                    if (++runs > 4096) throw new NotSupportedException("Text ink inspection exceeds its 4096-run work limit.");
                    if (original.Color.A == 0 || original.Width <= 0D || original.Height <= 0D) continue;
                    if (original.TextAdvanceWidth is not double measuredAdvance || original.Text.IndexOfAny(new[] { '\r', '\n' }) >= 0) { NotInspected(original, "Text without a single-line positioned advance cannot establish ink geometry."); continue; }
                    HtmlRenderText text = original.ResolveBaselineForPainting();
                    string value = text.BidiVisualOrderResolved ? "\u202D" + text.Text + "\u202C" : text.Text;
                    double advance = text.TextPaintWidth ?? (measuredAdvance > 0D ? measuredAdvance : text.Width);
                    double sourceSize = Math.Max(1D, original.Font.Size);
                    var ink = measurement.MeasurePositionedTextBounds(value, text.X, text.Y, text.Width, text.Height,
                        Math.Max(.1D, sourceSize * original.BaselineScale), text.Font, advance, text.Alignment, text.FeatureSettings, text.FontPalette,
                        sourceSize, text.UnderlineStyle, text.StrikethroughStyle, inkOnly: true,
                        inkTransform: transform, color: text.Color, decorationColor: text.DecorationColor);
                    if (!ink.IsMeasured || (ink.HasInk && (!Finite(ink.Left) || !Finite(ink.Top) || !Finite(ink.Right) || !Finite(ink.Bottom)))) { NotInspected(text, "A text outline was unavailable; fallback box estimates cannot establish glyph ink."); continue; }
                    if (ink.HasInk && (ink.Left < -.01D || ink.Top < -.01D || ink.Right > width + .01D || ink.Bottom > height + .01D))
                        diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Html", isRegion ? HtmlRenderDiagnosticCodes.TextInkOutsideRegion : HtmlRenderDiagnosticCodes.TextInkOutsideCanvas,
                            "Measured text paint bounds extend outside the " + (isRegion ? "region border box" : "declared canvas") + ". Decorations use conservative stroke bounds.",
                            HtmlDiagnosticSeverity.Warning, text.Source,
                            string.Format(CultureInfo.InvariantCulture, isRegion ? "left={0};top={1};right={2};bottom={3};regionWidth={4};regionHeight={5}" :
                                    "left={0};top={1};right={2};bottom={3};canvasWidth={4};canvasHeight={5}",
                                ink.Left, ink.Top, ink.Right, ink.Bottom, width, height)));
                } else Visit(InspectionChildren(visual), transform);
                if (diagnostics.Count > 4096) throw new NotSupportedException("Text ink inspection exceeds its diagnostic limit.");
            }
        }

        bool ContainsText(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (var visual in visuals) {
                token.ThrowIfCancellationRequested();
                if (visual is HtmlRenderText || visual is HtmlRenderDrawing || ContainsText(InspectionChildren(visual))) return true;
            }
            return false;
        }
        static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
        void NotInspected(HtmlRenderVisual visual, string message) => diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Html",
            HtmlRenderDiagnosticCodes.TextInkNotInspected, message, HtmlDiagnosticSeverity.Warning, visual.Source));
    }
}
