using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private const int MaximumEstimatedMetricRequests = 4096;
    private const int MaximumEstimatedMetricCharacters = 65536;
    private const int MaximumMetricDiagnosticIdentityCharacters = 256;
    private readonly HashSet<(string Family, OfficeFontFaceDescriptor Descriptor)> _estimatedMetricRequests = new();
    private int _estimatedMetricCharacters;
    private bool _estimatedMetricRequestsSaturated;

    /// <summary>
    /// Records only the branch that actually uses OfficeTextMeasurer. It retains no source
    /// text and adds no diagnostics while layout, font discovery or print fitting is provisional.
    /// </summary>
    private void RecordEstimatedTextMetrics(string text, OfficeFontInfo font, OfficeFontFaceDescriptor descriptor) {
        if (font.Size <= 0D || !HasPaintableMetricText(text)) return;
        RecordEstimatedTextMetricRequest(font.FamilyName, descriptor);
    }

    private void RecordEstimatedTextMetricRequest(string family, OfficeFontFaceDescriptor descriptor) {
        var key = (family, descriptor);
        if (_estimatedMetricRequests.Contains(key)) return;
        if (_estimatedMetricRequests.Count >= MaximumEstimatedMetricRequests
            || family.Length > MaximumEstimatedMetricCharacters - _estimatedMetricCharacters) {
            _estimatedMetricRequestsSaturated = true;
            return;
        }
        _estimatedMetricRequests.Add(key);
        _estimatedMetricCharacters += family.Length;
    }

    /// <summary>
    /// Carries actual child estimates only after the frame owner accepts its paint. The outer
    /// completed scene supplies final source attribution, rather than promoting a child's
    /// provisional warning snapshot into shared diagnostics during pagination or print fitting.
    /// </summary>
    private void AcceptEstimatedTextMetrics(HtmlRenderLayoutEngine child) {
        _estimatedMetricRequestsSaturated |= child._estimatedMetricRequestsSaturated;
        foreach (var request in child._estimatedMetricRequests) {
            CheckCancellation();
            ChargeLayoutOperations(1L, "frame font metric attribution");
            RecordEstimatedTextMetricRequest(request.Family, request.Descriptor);
        }
    }

    /// <summary>
    /// Attributes actual raw width estimates to paintable text in the completed layout.
    /// Exact retained shaping and covering scoped programs remain authoritative. Target metric
    /// delegates own their separate contract and are never reinvoked for diagnostic discovery.
    /// </summary>
    private HtmlRenderDocument ReportEstimatedTextMetrics(HtmlRenderDocument rendered) {
        // This report is deliberately raw-only. An internal target metric delegate remains
        // the reporting owner even when it declines an individual request. Covering that
        // target contract needs retained per-request provenance, not another delegate/provider call.
        if ((_estimatedMetricRequests.Count == 0 && !_estimatedMetricRequestsSaturated)
            || _options.FallbackTextMeasurement != null) return rendered;

        OfficeFontFaceCollection fonts = rendered.Fonts;
        var diagnostics = new List<HtmlDiagnostic>();
        var reported = new HashSet<(string Source, string Family, OfficeFontFaceDescriptor Descriptor)>();
        int reportedCharacters = 0;
        bool reportSaturated = false;
        var pending = new Stack<HtmlRenderVisual>();
        foreach (HtmlRenderPage page in rendered.Pages) {
            for (int index = page.Scene.Count - 1; index >= 0; index--) pending.Push(page.Scene[index]);
            while (pending.Count > 0) {
                CheckCancellation();
                ChargeLayoutOperations(1L, "font metric attribution");
                HtmlRenderVisual visual = pending.Pop();
                if (visual is HtmlRenderEffectGroup { Opacity: 0D }) continue;
                if (visual.PaintChildren is { } children) {
                    for (int index = children.Count - 1; index >= 0; index--) pending.Push(children[index]);
                }
                if (visual is not HtmlRenderText text || text.Font.Size <= 0D || text.Color.A == 0
                    || !HasPaintableMetricText(text.Text) || text.PreparedShaping != null) continue;
                if (!_estimatedMetricRequests.Contains((text.Font.FamilyName, text.FontDescriptor))
                    && !_estimatedMetricRequestsSaturated) continue;
                ChargeLayoutOperations(text.Text.Length, "font metric attribution text");
                bool covered = true;
                foreach (OfficeFontFallbackRun run in fonts.PlanFallbackRuns(text.Text, text.Font.FamilyName, text.FontDescriptor)) {
                    CheckCancellation();
                    if (!fonts.TryResolveFaceForText(run.Text, run.FamilyName, text.FontDescriptor,
                            text.Font.Size, out _)) {
                        covered = false;
                        break;
                    }
                }
                if (covered) continue;

                string source = BoundMetricIdentity(string.IsNullOrWhiteSpace(text.Source) ? "html:text" : text.Source!);
                string family = BoundMetricIdentity(text.Font.FamilyName);
                var key = (source, family, text.FontDescriptor);
                if (reported.Contains(key)) continue;
                int keyCharacters = source.Length + family.Length;
                if (reported.Count >= MaximumEstimatedMetricRequests
                    || keyCharacters > MaximumEstimatedMetricCharacters - reportedCharacters) {
                    reportSaturated = true;
                    continue;
                }
                reported.Add(key);
                reportedCharacters += keyCharacters;
                diagnostics.Add(new HtmlDiagnostic(ComponentName, HtmlRenderDiagnosticCodes.FontMetricsEstimated,
                    "Retained text used estimated widths without a selected font program.",
                    HtmlDiagnosticSeverity.Warning, source,
                    "estimator=OfficeTextMeasurer;axes=width;family=" + family
                        + ";weight=" + text.FontDescriptor.Weight
                        + ";stretch=" + text.FontDescriptor.StretchPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        + ";slant=" + text.FontDescriptor.Slant
                        + ";angle=" + text.FontDescriptor.ObliqueAngleDegrees.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    OfficeConversionLossKind.Approximation));
            }
        }
        if (reportSaturated) {
            diagnostics.Add(new HtmlDiagnostic(ComponentName, HtmlRenderDiagnosticCodes.FontMetricsEstimated,
                "Additional retained text used estimated widths after the bounded attribution limit was reached.",
                HtmlDiagnosticSeverity.Warning, "html:text", "estimator=OfficeTextMeasurer;axes=width;attribution=limited",
                OfficeConversionLossKind.Approximation));
        }
        return diagnostics.Count == 0 ? rendered : rendered.WithAdditionalDiagnostics(diagnostics);
    }

    private static string BoundMetricIdentity(string value) => value.Length <= MaximumMetricDiagnosticIdentityCharacters
        ? value : value.Substring(0, MaximumMetricDiagnosticIdentityCharacters);

    private static bool HasPaintableMetricText(string text) {
        for (int index = 0; index < text.Length; index++) {
            int scalar = text[index];
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                scalar = char.ConvertToUtf32(text[index], text[++index]);
            }
            if (!char.IsWhiteSpace(text[index]) && !OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) return true;
        }
        return false;
    }
}
