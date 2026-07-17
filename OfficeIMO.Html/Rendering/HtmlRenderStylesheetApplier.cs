using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlRenderStylesheetApplier {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static void Apply(IHtmlDocument document, HtmlRenderResourceSet resources, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        var reportedCycles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement link in document.QuerySelectorAll("link[href]")) {
            if (!IsStylesheetLink(link)) {
                continue;
            }

            string source = link.GetAttribute("href") ?? string.Empty;
            if (!resources.TryGet(source, null, out HtmlResolvedResource resource)) {
                continue;
            }

            if (!HtmlRenderStylesheetText.TryDecode(resource.Bytes, out string css)) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported,
                    "A resolved stylesheet could not be decoded as UTF-8 or BOM-declared UTF-16 CSS text.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    resource.ContentType);
                continue;
            }

            if (resources.TryGetResolvedSource(source, null, out string resolvedSource)
                && Uri.TryCreate(resolvedSource, UriKind.Absolute, out Uri? stylesheetUri)) {
                css = ExpandImports(
                    css,
                    stylesheetUri,
                    resources,
                    options,
                    diagnostics,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                    reportedCycles);
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
        HtmlRenderResourceSet resources,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        HashSet<string> activeStylesheets,
        HashSet<string> reportedCycles) {
        string currentKey = stylesheetUri.AbsoluteUri;
        if (!activeStylesheets.Add(currentKey)) {
            ReportCycle(diagnostics, reportedCycles, currentKey);
            return string.Empty;
        }

        var resourceOptions = new HtmlResourcePipelineOptions {
            UrlPolicy = options.GetResourceUrlPolicy().Clone(),
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
                && resources.TryGet(reference.Source, reference.ResolvedSource, out HtmlResolvedResource importedResource)
                && Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? importedUri)) {
                if (activeStylesheets.Contains(importedUri.AbsoluteUri)) {
                    ReportCycle(diagnostics, reportedCycles, importedUri.AbsoluteUri);
                } else if (HtmlRenderStylesheetText.TryDecode(importedResource.Bytes, out string importedCss)) {
                    replacement = ExpandImports(importedCss, importedUri, resources, options, diagnostics, activeStylesheets, reportedCycles);
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
