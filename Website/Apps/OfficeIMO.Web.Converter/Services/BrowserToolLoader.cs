using Microsoft.AspNetCore.Components.WebAssembly.Services;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>Loads format engines only when a browser tool needs them; successful loads are reused.</summary>
public sealed class BrowserToolLoader(LazyAssemblyLoader assemblies) {
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly HashSet<string> _loaded = new(StringComparer.Ordinal);

    internal static IReadOnlyList<string> RequiredAssemblies(string workspace, string route) {
        if (workspace == "provenance") {
            // The shared provenance workflow supports every Office container in its catalog.
            return ["OfficeIMO.Word.wasm", "OfficeIMO.Excel.wasm", "OfficeIMO.PowerPoint.wasm",
                "OfficeIMO.Visio.wasm", "OfficeIMO.Workflows.wasm", "OfficeIMO.Web.Fonts.wasm",
                "OfficeIMO.Pdf.wasm", "OfficeIMO.Html.wasm", "OfficeIMO.Markdown.wasm",
                "OfficeIMO.Project.wasm", "OfficeIMO.Reader.Core.wasm", "DocumentFormat.OpenXml.wasm",
                "AngleSharp.wasm", "AngleSharp.Css.wasm"];
        }
        if (workspace == "pdf") return ["OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"];
        return route switch {
            "docx-pdf" or "pdf-docx" => ["OfficeIMO.Word.wasm", "DocumentFormat.OpenXml.wasm", "OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"],
            "xlsx-pdf" or "pdf-xlsx" => ["OfficeIMO.Excel.wasm", "DocumentFormat.OpenXml.wasm", "OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"],
            "pptx-pdf" or "pdf-pptx" => ["OfficeIMO.PowerPoint.wasm", "DocumentFormat.OpenXml.wasm", "OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"],
            "markdown-docx" => ["OfficeIMO.Word.wasm", "DocumentFormat.OpenXml.wasm", "OfficeIMO.Markdown.wasm", "OfficeIMO.Html.wasm", "AngleSharp.wasm", "AngleSharp.Css.wasm"],
            "html-pdf" => ["OfficeIMO.Html.wasm", "AngleSharp.wasm", "AngleSharp.Css.wasm", "OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"],
            "pdf-html" => ["OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm", "OfficeIMO.Html.wasm", "AngleSharp.wasm", "AngleSharp.Css.wasm"],
            "pdf-png" => ["OfficeIMO.Pdf.wasm", "OfficeIMO.Web.Fonts.wasm"],
            "markdown-html" => ["OfficeIMO.Markdown.wasm"],
            "html-markdown" => ["OfficeIMO.Markdown.wasm", "OfficeIMO.Html.wasm", "AngleSharp.wasm", "AngleSharp.Css.wasm"],
            _ => []
        };
    }

    /// <summary>Awaits the selected tool's dependencies. Download failures are surfaced to the workspace.</summary>
    public async Task LoadAsync(string workspace, string route) {
        await _gate.WaitAsync();
        try {
            string[] pending = RequiredAssemblies(workspace, route).Where(name => !_loaded.Contains(name)).ToArray();
            if (pending.Length == 0) return;
            await assemblies.LoadAssembliesAsync(pending);
            _loaded.UnionWith(pending);
        } finally {
            _gate.Release();
        }
    }
}
