using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using System.Globalization;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    private readonly HtmlCssMediaContext _sourceMediaContext;
    private bool _useCanonicalLayerCascade;

    internal HtmlToWordConverter(HtmlCssMediaContext sourceMediaContext = HtmlCssMediaContext.Screen) {
        _sourceMediaContext = sourceMediaContext;
    }

    /// <summary>Projects accepted layered author CSS into the operation-owned import DOM.</summary>
    private void MaterializeLayeredCascade(IHtmlDocument document) {
        if (!_useCanonicalLayerCascade) return;

        // Register the owned fallback layer before every admitted author layer, including
        // configured stylesheets. Admission order and the relative author order stay intact.
        var themeSources = new HashSet<string>(document.QuerySelectorAll("style[data-officeimo-theme]").Select(style => style.TextContent), StringComparer.Ordinal);
        if (themeSources.Count > 0) {
            var ordered = _pageStylesheets.OrderBy(stylesheet => themeSources.Contains(stylesheet.Css) ? 0 : 1).ToArray();
            _pageStylesheets.Clear();
            _pageStylesheets.AddRange(ordered);
        }

        // The importer already admitted these sources in order, with its resource and byte limits.
        // Reuse them without fetching again, and remove the original nodes to avoid a second cascade.
        foreach (IElement style in document.QuerySelectorAll("style").ToArray()) style.Remove();
        foreach (var stylesheet in _pageStylesheets) {
            _cancellationToken.ThrowIfCancellationRequested();
            IElement style = document.CreateElement("style");
            style.TextContent = stylesheet.Css;
            if (!string.IsNullOrWhiteSpace(stylesheet.Media)) style.SetAttribute("media", stylesheet.Media);
            document.Head!.AppendChild(style);
        }

        IReadOnlyDictionary<IElement, HtmlComputedStyle> computed;
        try {
            computed = HtmlComputedStyleEngine.Compute(document, _sourceMediaContext, _options.Limits);
        } catch (HtmlDomLimitException exception) {
            ThrowLimitExceeded(_options, exception.Code, exception.Message, exception.LimitSource, exception.Actual, exception.Limit);
            throw;
        }

        var authoredProperties = new Dictionary<IElement, HashSet<string>>();
        string[] resetCandidates = _supportedCssDiagnosticProperties
            .Concat(_blockSpacingCssDiagnosticProperties).Concat(_tableSpacingCssDiagnosticProperties)
            .Concat(HtmlCssPropertyCatalog.All.Select(property => property.Name)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        foreach (IElement element in document.All) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!computed.TryGetValue(element, out HtmlComputedStyle? style)) continue;
            HashSet<string>? parentProperties = element.ParentElement != null && authoredProperties.TryGetValue(element.ParentElement, out var inherited)
                ? inherited : null;
            var projected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var authored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in style.Properties) {
                bool fromAuthor = style.IsSpecifiedValue(property.Key) || style.IsResetValue(property.Key) ||
                    style.IsInheritedValue(property.Key) && parentProperties?.Contains(property.Key) == true;
                if (!fromAuthor) continue;
                authored.Add(property.Key);
                if (property.Key.StartsWith("--", StringComparison.Ordinal)) continue;
                projected[property.Key] = property.Value;
                if (style.IsImplicitlyInheritedValue(property.Key)) {
                    if (!_injectedInheritedCssProperties.TryGetValue(element, out HashSet<string>? injected)) {
                        injected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        _injectedInheritedCssProperties[element] = injected;
                    }
                    injected.Add(property.Key);
                }
            }
            foreach (string property in resetCandidates) {
                if (style.IsResetValue(property) && !projected.ContainsKey(property)) {
                    projected[property] = "initial";
                    authored.Add(property);
                }
            }
            if (authored.Contains("font-size") && style.ResolvedFontSizePoints is double size) {
                // Relative sizes have already been resolved against the original ancestor chain.
                projected["font-size"] = size.ToString("R", CultureInfo.InvariantCulture) + "pt";
                _computedFontSizePixels[element] = size * 96D / 72D;
            }
            authoredProperties[element] = authored;
            if (projected.Count == 0) continue;
            ReportUnsupportedInlineCssDiagnostics(element, authored);
            LowerMappedLayerValues(element, projected);
            var css = new System.Text.StringBuilder();
            // Replay expanded longhands after their shorthands, preserving the owner's winning values.
            foreach (var property in projected.OrderBy(property => property.Key.Count(character => character == '-'))) {
                css.Append(property.Key).Append(':').Append(property.Value).Append(';');
            }
            element.SetAttribute("style", css.ToString());
        }
        // These inline values are the canonical cascade result, including important/layer ordering.
        // The legacy importer must not replay its flat subset over that result.
        _cssRules.Clear();
    }
}
