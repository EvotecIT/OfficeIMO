using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlRenderStylesheetApplier {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static HtmlCssByteBudget CreateBudget(IHtmlDocument document, HtmlConversionLimits limits) {
        var budget = new HtmlCssByteBudget(limits);
        foreach (IElement style in document.QuerySelectorAll("style")) {
            budget.ReserveOrThrow(style.TextContent ?? string.Empty);
        }

        return budget;
    }

    internal static void Apply(
        IHtmlDocument document,
        HtmlResourceSession resources,
        HtmlRenderOptions options,
        HtmlCssByteBudget cssBudget,
        HtmlDiagnosticReport diagnostics) {
        var reportedCycles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement link in document.QuerySelectorAll("link[href]")) {
            if (!IsStylesheetLink(link)) {
                continue;
            }

            string source = link.GetAttribute("href") ?? string.Empty;
            string? resolvedSource = resources.TryGetResolvedSource(source, null, out string resolved)
                ? resolved
                : null;
            if (resources.WasStylesheetRejected(source, resolvedSource)) {
                continue;
            }

            if (!resources.TryGet(source, null, out HtmlResolvedResource resource)) {
                continue;
            }

            if (!HtmlRenderStylesheetText.TryDecode(resource.EncodedBytes, out string css)) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported,
                    "A resolved stylesheet could not be decoded as UTF-8 or BOM-declared UTF-16 CSS text.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    resource.ContentType);
                continue;
            }

            if (!resources.WasStylesheetBudgeted(source, resolvedSource)
                && !TryReserveCss(cssBudget, css, source, diagnostics)) {
                continue;
            }

            if (resolvedSource != null
                && Uri.TryCreate(resolvedSource, UriKind.Absolute, out Uri? stylesheetUri)) {
                css = ExpandImports(
                    css,
                    stylesheetUri,
                    resources,
                    options,
                    diagnostics,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                    reportedCycles,
                    cssBudget);
            }

            if (HtmlResourcePipeline.HasStylesheetUrlResources(css)) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending,
                    "The external stylesheet was applied, but its URL resources are not active in the current paint model.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    resource.ContentType);
            }

            IElement style = document.CreateElement("style");
            style.TextContent = css;
            style.SetAttribute("data-officeimo-source", source);
            string media = link.GetAttribute("media") ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(media)) {
                style.SetAttribute("media", media);
            }

            INode? parent = link.Parent;
            if (parent == null) {
                continue;
            }

            INode? next = link.NextSibling;
            if (next == null) {
                parent.AppendChild(style);
            } else {
                parent.InsertBefore(style, next);
            }
        }
    }

    private static string ExpandImports(
        string css,
        Uri stylesheetUri,
        HtmlResourceSession resources,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        HashSet<string> activeStylesheets,
        HashSet<string> reportedCycles,
        HtmlCssByteBudget cssBudget) {
        string currentKey = stylesheetUri.AbsoluteUri;
        if (!activeStylesheets.Add(currentKey)) {
            ReportCycle(diagnostics, reportedCycles, currentKey);
            return string.Empty;
        }

        var resourceOptions = new HtmlResourcePipelineOptions {
            ResourceUrlPolicy = options.GetResourceUrlPolicy().Clone(),
            MaxResponsiveImageCandidates = options.ResponsiveImageCandidateLimit,
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D
        };
        HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(css, stylesheetUri, resourceOptions);
        var builder = new System.Text.StringBuilder(analysis.Css);
        for (int index = analysis.Imports.Count - 1; index >= 0; index--) {
            HtmlExternalStylesheetImport import = analysis.Imports[index];
            string replacement = string.Empty;
            HtmlResourceReference reference = import.Reference;
            if (import.IsApplicable
                && reference.IsAllowed
                && !resources.WasStylesheetRejected(reference.Source, reference.ResolvedSource)
                && resources.TryGet(reference.Source, reference.ResolvedSource, out HtmlResolvedResource importedResource)
                && Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? importedUri)) {
                if (activeStylesheets.Contains(importedUri.AbsoluteUri)) {
                    ReportCycle(diagnostics, reportedCycles, importedUri.AbsoluteUri);
                } else if (HtmlRenderStylesheetText.TryDecode(importedResource.EncodedBytes, out string importedCss)) {
                    if (resources.WasStylesheetBudgeted(reference.Source, reference.ResolvedSource)
                        || TryReserveCss(cssBudget, importedCss, reference.Source, diagnostics)) {
                        replacement = ExpandImports(importedCss, importedUri, resources, options, diagnostics, activeStylesheets, reportedCycles, cssBudget);
                    }
                } else {
                    diagnostics.Add(
                        ComponentName,
                        HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported,
                        "An imported stylesheet could not be decoded as UTF-8 or BOM-declared UTF-16 CSS text.",
                        HtmlDiagnosticSeverity.Warning,
                        reference.Source,
                        importedResource.ContentType);
                }
            }

            builder.Remove(import.Start, import.End - import.Start);
            builder.Insert(import.Start, replacement);
        }

        activeStylesheets.Remove(currentKey);
        return HtmlResourcePipeline.RebaseExternalStylesheetUrls(
            builder.ToString(),
            stylesheetUri,
            options.GetResourceUrlPolicy());
    }

    internal static bool TryReserveCss(
        HtmlCssByteBudget budget,
        string css,
        string source,
        HtmlDiagnosticReport diagnostics) {
        if (budget.TryReserve(css, out HtmlDomLimitException? exception)) return true;
        diagnostics.Add(
            ComponentName,
            exception!.Code,
            "A resolved stylesheet was omitted because it exceeded the shared CSS byte budget.",
            HtmlDiagnosticSeverity.Error,
            source,
            exception.LimitSource + ": " + exception.Detail,
            HtmlConversionLossKind.Omission);
        return false;
    }

    private static void ReportCycle(HtmlDiagnosticReport diagnostics, HashSet<string> reportedCycles, string source) {
        if (!reportedCycles.Add(source)) {
            return;
        }

        diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.StylesheetImportCycle,
            "A recursive stylesheet import cycle was suppressed.",
            HtmlDiagnosticSeverity.Warning,
            source);
    }

    private static bool IsStylesheetLink(IElement link) {
        string rel = link.GetAttribute("rel") ?? string.Empty;
        foreach (string token in rel.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }
}
