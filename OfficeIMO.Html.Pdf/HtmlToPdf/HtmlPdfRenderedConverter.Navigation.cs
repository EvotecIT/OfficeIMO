using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>Admits annotations only when their targets can be represented in the selected PDF.</summary>
    internal sealed class LinkTargets {
        private readonly HashSet<string> _destinations = new(StringComparer.Ordinal);
        private readonly HashSet<string> _reported = new(StringComparer.Ordinal);
        private readonly bool _hasUriBase;
        private readonly bool _allowDeferredDestinations;
        private readonly PdfCore.PdfConversionReport _report;

        internal LinkTargets(HtmlRenderDocument document, PdfCore.PdfCreateOptions options,
            PdfCore.PdfConversionReport report, CancellationToken cancellationToken, bool allowDeferredDestinations = false) {
            _allowDeferredDestinations = allowDeferredDestinations;
            _hasUriBase = !string.IsNullOrWhiteSpace(options.CatalogUriBase);
            _report = report;
            if (_allowDeferredDestinations) return;
            foreach (HtmlRenderVisual visual in EnumerateVisuals(document.Pages.SelectMany(page => page.Scene))) {
                cancellationToken.ThrowIfCancellationRequested();
                if (visual is HtmlRenderNamedDestination destination) _destinations.Add(destination.Name);
            }
        }

        internal bool Include(string? uri, string? source) {
            if (uri == null) return true;
            bool fragment = IsFragmentLink(uri);
            if (fragment ? _allowDeferredDestinations || _destinations.Contains(uri.Substring(1))
                : _hasUriBase || Uri.TryCreate(uri, UriKind.Absolute, out _)) return true;
            if (_reported.Add(uri)) {
                _report.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Html.Pdf", HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable,
                    source ?? "html-link",
                    fragment
                        ? "A document-internal hyperlink target is outside the selected output pages and the link annotation was omitted."
                        : "A relative hyperlink has no base URI and the link annotation was omitted.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission,
                    details: new Dictionary<string, string> { ["Target"] = uri }));
            }
            return false;
        }
    }

    /// <summary>Retains navigation metadata when an empty paint clip culls a subtree.</summary>
    private static void AddClippedNamedDestinations(
        PdfCore.PdfPageCanvas canvas,
        IEnumerable<HtmlRenderVisual> visuals,
        OfficeTransform transform,
        CancellationToken cancellationToken) {
        foreach ((string name, OfficePoint point) in EnumerateNamedDestinationPositions(visuals, transform, cancellationToken)) {
            canvas.NamedDestination(MapNamedDestination(name),
                global::OfficeIMO.OfficeLength.Points(point.X * PointsPerCssPixel), global::OfficeIMO.OfficeLength.Points(point.Y * PointsPerCssPixel));
        }
    }

    private static IEnumerable<(string Name, OfficePoint Point)> EnumerateNamedDestinationPositions(
        IEnumerable<HtmlRenderVisual> visuals, OfficeTransform transform, CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderNamedDestination destination) {
                yield return (destination.Name, transform.TransformPoint(new OfficePoint(destination.X, destination.Y)));
                continue;
            }
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup path => path.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderLayoutRegion layout => layout.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children == null) continue;
            OfficeTransform childTransform = visual is HtmlRenderEffectGroup group ? group.Transform.Then(transform) : transform;
            foreach (var target in EnumerateNamedDestinationPositions(children, childTransform, cancellationToken))
                yield return target;
        }
    }
}
