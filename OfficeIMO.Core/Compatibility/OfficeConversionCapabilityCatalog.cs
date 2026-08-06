using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace OfficeIMO;

/// <summary>Describes how a conversion route accepts its source content.</summary>
public enum OfficeConversionInputKind {
    /// <summary>The route accepts a document file or stream.</summary>
    File,
    /// <summary>The route accepts source text.</summary>
    Text
}

/// <summary>Describes the primary fidelity contract of a conversion route.</summary>
public enum OfficeConversionFidelityKind {
    /// <summary>The destination preserves editable semantic structure where representable.</summary>
    Editable,
    /// <summary>The destination prioritizes fixed-layout visual output.</summary>
    FixedLayout,
    /// <summary>The destination projects normalized semantic content.</summary>
    Semantic
}

/// <summary>One package-neutral conversion route exposed by OfficeIMO.</summary>
public sealed class OfficeConversionCapability {
    /// <summary>Creates a conversion capability.</summary>
    public OfficeConversionCapability(
        string id,
        string source,
        string target,
        OfficeConversionInputKind inputKind,
        IEnumerable<string> sourceExtensions,
        string targetExtension,
        string packageId,
        string api,
        string description,
        OfficeConversionFidelityKind fidelity,
        string resultContract,
        bool browserAvailable = false,
        bool agentDiscoverable = true) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Route id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("Source label cannot be empty.", nameof(source));
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("Target label cannot be empty.", nameof(target));
        if (sourceExtensions == null) throw new ArgumentNullException(nameof(sourceExtensions));
        if (string.IsNullOrWhiteSpace(targetExtension)) throw new ArgumentException("Target extension cannot be empty.", nameof(targetExtension));
        if (string.IsNullOrWhiteSpace(packageId)) throw new ArgumentException("Package id cannot be empty.", nameof(packageId));
        if (string.IsNullOrWhiteSpace(api)) throw new ArgumentException("API cannot be empty.", nameof(api));
        if (string.IsNullOrWhiteSpace(description)) throw new ArgumentException("Description cannot be empty.", nameof(description));
        if (string.IsNullOrWhiteSpace(resultContract)) throw new ArgumentException("Result contract cannot be empty.", nameof(resultContract));

        string[] extensions = sourceExtensions
            .Select(NormalizeExtension)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (extensions.Length == 0) throw new ArgumentException("At least one source extension is required.", nameof(sourceExtensions));

        Id = id.Trim();
        Source = source.Trim();
        Target = target.Trim();
        InputKind = inputKind;
        SourceExtensions = new ReadOnlyCollection<string>(extensions);
        TargetExtension = NormalizeExtension(targetExtension);
        PackageId = packageId.Trim();
        Api = api.Trim();
        Description = description.Trim();
        Fidelity = fidelity;
        ResultContract = resultContract.Trim();
        BrowserAvailable = browserAvailable;
        AgentDiscoverable = agentDiscoverable;
    }

    /// <summary>Gets the stable route identifier.</summary>
    public string Id { get; }
    /// <summary>Gets the source format label.</summary>
    public string Source { get; }
    /// <summary>Gets the destination format label.</summary>
    public string Target { get; }
    /// <summary>Gets whether the route accepts a file or source text.</summary>
    public OfficeConversionInputKind InputKind { get; }
    /// <summary>Gets normalized accepted source extensions.</summary>
    public IReadOnlyList<string> SourceExtensions { get; }
    /// <summary>Gets the normalized destination extension.</summary>
    public string TargetExtension { get; }
    /// <summary>Gets the focused package that owns the adapter.</summary>
    public string PackageId { get; }
    /// <summary>Gets the representative public API.</summary>
    public string Api { get; }
    /// <summary>Gets the customer-facing route description.</summary>
    public string Description { get; }
    /// <summary>Gets the primary fidelity contract.</summary>
    public OfficeConversionFidelityKind Fidelity { get; }
    /// <summary>Gets the public result contract returned by the representative API.</summary>
    public string ResultContract { get; }
    /// <summary>Gets whether the route executes in the shipped browser converter.</summary>
    public bool BrowserAvailable { get; }
    /// <summary>Gets whether the route is advertised through agent capability discovery.</summary>
    public bool AgentDiscoverable { get; }

    private static string NormalizeExtension(string extension) {
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("Extensions cannot be empty.", nameof(extension));
        string value = extension.Trim();
        return (value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value).ToLowerInvariant();
    }
}

/// <summary>The shared OfficeIMO conversion route catalog used by packages, agents, and browser surfaces.</summary>
public static class OfficeConversionCapabilityCatalog {
    /// <summary>Gets the capability schema version.</summary>
    public const int SchemaVersion = 1;

    /// <summary>Gets all focused, public document-conversion routes in stable order.</summary>
    public static IReadOnlyList<OfficeConversionCapability> All { get; } =
        new ReadOnlyCollection<OfficeConversionCapability>(CreateRoutes());

    /// <summary>Gets the routes available in the shipped WebAssembly converter.</summary>
    public static IReadOnlyList<OfficeConversionCapability> BrowserRoutes { get; } =
        new ReadOnlyCollection<OfficeConversionCapability>(All.Where(static route => route.BrowserAvailable).ToArray());

    /// <summary>Gets the routes advertised by agent discovery.</summary>
    public static IReadOnlyList<OfficeConversionCapability> AgentRoutes { get; } =
        new ReadOnlyCollection<OfficeConversionCapability>(All.Where(static route => route.AgentDiscoverable).ToArray());

    /// <summary>Finds a route by its stable id.</summary>
    public static OfficeConversionCapability? Find(string? id) =>
        string.IsNullOrWhiteSpace(id)
            ? null
            : All.FirstOrDefault(route => string.Equals(route.Id, id!.Trim(), StringComparison.OrdinalIgnoreCase));

    /// <summary>Returns routes that accept the supplied source extension.</summary>
    public static IReadOnlyList<OfficeConversionCapability> FindBySourceExtension(string extension) {
        string normalized = NormalizeExtension(extension);
        return All.Where(route => route.SourceExtensions.Contains(normalized, StringComparer.OrdinalIgnoreCase)).ToArray();
    }

    /// <summary>Formats the conversion routes as a deterministic Markdown reference.</summary>
    public static string ToMarkdown() {
        var markdown = new StringBuilder();
        markdown.AppendLine("# OfficeIMO conversion routes");
        markdown.AppendLine();
        markdown.AppendLine("Use this table to find the focused package, representative API, fidelity model, and result type for a source-to-target conversion.");
        markdown.AppendLine();
        markdown.Append("Schema version: ").Append(SchemaVersion).AppendLine();
        markdown.AppendLine();
        markdown.AppendLine("| Route | Source | Target | Package | Fidelity | Browser | API | Result type | What it does |");
        markdown.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        foreach (OfficeConversionCapability route in All) {
            markdown.Append("| ").Append(EscapeMarkdown(route.Id))
                .Append(" | ").Append(EscapeMarkdown(route.Source))
                .Append(" | ").Append(EscapeMarkdown(route.Target))
                .Append(" | ").Append(EscapeMarkdown(route.PackageId))
                .Append(" | ").Append(route.Fidelity)
                .Append(" | ").Append(route.BrowserAvailable ? "Yes" : "No")
                .Append(" | `").Append(EscapeMarkdown(route.Api)).Append("`")
                .Append(" | ").Append(EscapeMarkdown(route.ResultContract))
                .Append(" | ").Append(EscapeMarkdown(route.Description)).AppendLine(" |");
        }
        return markdown.ToString().Replace("\r\n", "\n").Replace("\r", "\n");
    }

    /// <summary>Serializes the shared route contract as deterministic JSON.</summary>
    public static string ToJson(bool indented = true) {
        string newline = indented ? "\n" : string.Empty;
        string i1 = indented ? "  " : string.Empty;
        string i2 = indented ? "    " : string.Empty;
        string i3 = indented ? "      " : string.Empty;
        var json = new StringBuilder();
        json.Append('{').Append(newline)
            .Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(newline)
            .Append(i1).Append("\"routes\": [").Append(newline);
        for (int index = 0; index < All.Count; index++) {
            OfficeConversionCapability route = All[index];
            json.Append(i2).Append('{').Append(newline)
                .Append(i3).Append("\"id\":\"").Append(EscapeJson(route.Id)).Append("\",").Append(newline)
                .Append(i3).Append("\"source\":\"").Append(EscapeJson(route.Source)).Append("\",").Append(newline)
                .Append(i3).Append("\"target\":\"").Append(EscapeJson(route.Target)).Append("\",").Append(newline)
                .Append(i3).Append("\"inputKind\":\"").Append(route.InputKind).Append("\",").Append(newline)
                .Append(i3).Append("\"sourceExtensions\":[");
            for (int extensionIndex = 0; extensionIndex < route.SourceExtensions.Count; extensionIndex++) {
                if (extensionIndex > 0) json.Append(',');
                if (indented) json.Append(' ');
                json.Append('"').Append(EscapeJson(route.SourceExtensions[extensionIndex])).Append('"');
            }
            if (indented && route.SourceExtensions.Count > 0) json.Append(' ');
            json.Append("],").Append(newline)
                .Append(i3).Append("\"targetExtension\":\"").Append(EscapeJson(route.TargetExtension)).Append("\",").Append(newline)
                .Append(i3).Append("\"packageId\":\"").Append(EscapeJson(route.PackageId)).Append("\",").Append(newline)
                .Append(i3).Append("\"api\":\"").Append(EscapeJson(route.Api)).Append("\",").Append(newline)
                .Append(i3).Append("\"description\":\"").Append(EscapeJson(route.Description)).Append("\",").Append(newline)
                .Append(i3).Append("\"fidelity\":\"").Append(route.Fidelity).Append("\",").Append(newline)
                .Append(i3).Append("\"resultContract\":\"").Append(EscapeJson(route.ResultContract)).Append("\",").Append(newline)
                .Append(i3).Append("\"browserAvailable\":").Append(route.BrowserAvailable ? "true" : "false").Append(',').Append(newline)
                .Append(i3).Append("\"agentDiscoverable\":").Append(route.AgentDiscoverable ? "true" : "false").Append(newline)
                .Append(i2).Append('}');
            if (index + 1 < All.Count) json.Append(',');
            json.Append(newline);
        }
        json.Append(i1).Append(']').Append(newline).Append('}');
        return json.ToString();
    }

    private static OfficeConversionCapability[] CreateRoutes() => new[] {
        Route("docx-pdf", "DOCX", "PDF", OfficeConversionInputKind.File, new[] { ".docx" }, ".pdf", "OfficeIMO.Word.Pdf", "WordDocument.Load(stream).ToPdfDocumentResult(options)", "Convert a Word document into a fixed-layout PDF with diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult", browser: true),
        Route("xlsx-pdf", "XLSX", "PDF", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".pdf", "OfficeIMO.Excel.Pdf", "ExcelDocument.Load(stream).ToPdfDocumentResult(options)", "Render workbook sheets with layout and conversion diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult", browser: true),
        Route("pptx-pdf", "PPTX", "PDF", OfficeConversionInputKind.File, new[] { ".pptx" }, ".pdf", "OfficeIMO.PowerPoint.Pdf", "PowerPointPresentation.Load(stream).ToPdfDocumentResult(options)", "Render presentation slides into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult", browser: true),
        Route("html-pdf", "HTML", "PDF", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".pdf", "OfficeIMO.Html.Pdf", "HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options)", "Render bounded HTML and CSS into a tagged PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult", browser: true),
        Route("markdown-html", "Markdown", "HTML", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".html", "OfficeIMO.MarkdownRenderer", "MarkdownRenderer.RenderBodyHtml(markdown, options)", "Render typed Markdown through an explicit safe HTML profile.", OfficeConversionFidelityKind.Semantic, "string", browser: true),
        Route("html-markdown", "HTML", "Markdown", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".md", "OfficeIMO.Markdown.Html", "HtmlConversionDocument.Parse(html).ToMarkdownDocumentResult(options)", "Project HTML into portable Markdown with explicit resource and loss policy.", OfficeConversionFidelityKind.Semantic, "HtmlToMarkdownResult", browser: true),
        Route("markdown-docx", "Markdown", "DOCX", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".docx", "OfficeIMO.Word.Markdown", "MarkdownReader.Parse(markdown).ToWordDocumentResult(options)", "Create an editable Word document from typed Markdown.", OfficeConversionFidelityKind.Editable, "MarkdownToWordResult", browser: true),
        Route("docx-html", "DOCX", "HTML", OfficeConversionInputKind.File, new[] { ".docx" }, ".html", "OfficeIMO.Word.Html", "WordDocument.Load(stream).ToHtmlResult(options)", "Project a Word document into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "HtmlTextConversionResult"),
        Route("docx-markdown", "DOCX", "Markdown", OfficeConversionInputKind.File, new[] { ".docx" }, ".md", "OfficeIMO.Word.Markdown", "WordDocument.Load(stream).ToMarkdownDocumentResult(options)", "Convert a Word document into portable Markdown for editing, publishing, or version control.", OfficeConversionFidelityKind.Semantic, "WordToMarkdownResult"),
        Route("docx-odt", "DOCX", "ODT", OfficeConversionInputKind.File, new[] { ".docx" }, ".odt", "OfficeIMO.Word.OpenDocument", "WordDocument.Load(stream).ToOpenDocumentResult(options)", "Convert editable Word content to OpenDocument Text.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<OdtDocument>"),
        Route("odt-docx", "ODT", "DOCX", OfficeConversionInputKind.File, new[] { ".odt" }, ".docx", "OfficeIMO.Word.OpenDocument", "OdtDocument.Load(stream).ToWordDocumentResult(options)", "Convert OpenDocument Text into an editable Word document.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<WordDocument>"),
        Route("docx-rtf", "DOCX", "RTF", OfficeConversionInputKind.File, new[] { ".docx" }, ".rtf", "OfficeIMO.Word.Rtf", "WordDocument.Load(stream).ToRtfDocumentResult()", "Convert editable Word content to semantic RTF.", OfficeConversionFidelityKind.Editable, "RtfConversionResult<RtfDocument>"),
        Route("rtf-docx", "RTF", "DOCX", OfficeConversionInputKind.File, new[] { ".rtf" }, ".docx", "OfficeIMO.Word.Rtf", "RtfDocument.Load(stream, readOptions).ToWordDocumentResult(sourcePath)", "Convert bounded RTF into an editable Word document while retaining read and conversion diagnostics.", OfficeConversionFidelityKind.Editable, "RtfConversionResult<WordDocument>"),
        Route("xlsx-html", "XLSX", "HTML", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".html", "OfficeIMO.Excel.Html", "ExcelDocument.Load(stream).ToHtmlResult(options)", "Project workbook content into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "HtmlTextConversionResult"),
        Route("xlsx-ods", "XLSX", "ODS", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".ods", "OfficeIMO.Excel.OpenDocument", "ExcelDocument.Load(stream).ToOpenDocumentResult(options)", "Convert an editable workbook to OpenDocument Spreadsheet.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<OdsDocument>"),
        Route("ods-xlsx", "ODS", "XLSX", OfficeConversionInputKind.File, new[] { ".ods" }, ".xlsx", "OfficeIMO.Excel.OpenDocument", "OdsDocument.Load(stream).ToExcelDocumentResult(options)", "Convert OpenDocument Spreadsheet into an editable workbook.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<ExcelDocument>"),
        Route("pptx-html", "PPTX", "HTML", OfficeConversionInputKind.File, new[] { ".pptx" }, ".html", "OfficeIMO.PowerPoint.Html", "PowerPointPresentation.Load(stream).ToHtmlResult(options)", "Project presentation slides into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "PowerPointToHtmlResult"),
        Route("pptx-odp", "PPTX", "ODP", OfficeConversionInputKind.File, new[] { ".pptx" }, ".odp", "OfficeIMO.PowerPoint.OpenDocument", "PowerPointPresentation.Load(stream).ToOpenDocumentResult(options)", "Convert an editable presentation to OpenDocument Presentation.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<OdpPresentation>"),
        Route("odp-pptx", "ODP", "PPTX", OfficeConversionInputKind.File, new[] { ".odp" }, ".pptx", "OfficeIMO.PowerPoint.OpenDocument", "OdpPresentation.Load(stream).ToPowerPointPresentationResult(options)", "Convert OpenDocument Presentation into an editable presentation.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<PowerPointPresentation>"),
        Route("markdown-pdf", "Markdown", "PDF", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".pdf", "OfficeIMO.Markdown.Pdf", "MarkdownReader.Parse(markdown).ToPdfDocumentResult(options)", "Render Markdown into PDF with structured conversion warnings.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("rtf-markdown", "RTF", "Markdown", OfficeConversionInputKind.File, new[] { ".rtf" }, ".md", "OfficeIMO.Rtf.Markdown", "RtfDocument.Load(stream, readOptions).Document.ToMarkdownResult(options)", "Project bounded RTF into portable Markdown.", OfficeConversionFidelityKind.Semantic, "RtfConversionResult<string>"),
        Route("rtf-pdf", "RTF", "PDF", OfficeConversionInputKind.File, new[] { ".rtf" }, ".pdf", "OfficeIMO.Rtf.Pdf", "RtfDocument.Load(stream, readOptions).Document.ToPdfDocumentResult(options)", "Render bounded RTF into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("markdown-rtf", "Markdown", "RTF", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".rtf", "OfficeIMO.Rtf.Markdown", "MarkdownReader.Parse(markdown).ToRtfDocumentResult(options)", "Convert typed Markdown into semantic RTF with loss diagnostics.", OfficeConversionFidelityKind.Editable, "RtfConversionResult<RtfDocument>"),
        Route("html-docx", "HTML", "DOCX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".docx", "OfficeIMO.Word.Html", "HtmlConversionDocument.Parse(html).ToWordDocumentResult(options)", "Convert bounded HTML into an editable Word document.", OfficeConversionFidelityKind.Editable, "HtmlToWordResult"),
        Route("html-xlsx", "HTML", "XLSX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".xlsx", "OfficeIMO.Excel.Html", "HtmlConversionDocument.Parse(html).ToExcelDocumentResult(options)", "Convert bounded HTML tables and content into an editable workbook.", OfficeConversionFidelityKind.Editable, "HtmlToExcelResult"),
        Route("html-pptx", "HTML", "PPTX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".pptx", "OfficeIMO.PowerPoint.Html", "HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(options)", "Convert bounded HTML into an editable presentation.", OfficeConversionFidelityKind.Editable, "HtmlToPowerPointResult"),
        Route("html-rtf", "HTML", "RTF", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".rtf", "OfficeIMO.Html", "HtmlConversionDocument.Parse(html).ToRtfDocumentResult(options)", "Convert bounded HTML into semantic RTF.", OfficeConversionFidelityKind.Editable, "HtmlToRtfResult"),
        Route("rtf-html", "RTF", "HTML", OfficeConversionInputKind.File, new[] { ".rtf" }, ".html", "OfficeIMO.Html", "RtfDocument.Load(stream, readOptions).Document.ToHtmlResult(options)", "Render bounded RTF through an explicit safe HTML profile.", OfficeConversionFidelityKind.Semantic, "RtfToHtmlResult"),
        Route("asciidoc-markdown", "AsciiDoc", "Markdown", OfficeConversionInputKind.Text, new[] { ".adoc", ".asciidoc", ".txt" }, ".md", "OfficeIMO.AsciiDoc.Markdown", "AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult(options)", "Project AsciiDoc into typed Markdown with diagnostics.", OfficeConversionFidelityKind.Semantic, "AsciiDocToMarkdownResult"),
        Route("markdown-asciidoc", "Markdown", "AsciiDoc", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".adoc", "OfficeIMO.AsciiDoc.Markdown", "MarkdownReader.Parse(markdown).ToAsciiDocDocumentResult(options)", "Project typed Markdown into canonical AsciiDoc with diagnostics.", OfficeConversionFidelityKind.Semantic, "MarkdownToAsciiDocResult"),
        Route("asciidoc-pdf", "AsciiDoc", "PDF", OfficeConversionInputKind.Text, new[] { ".adoc", ".asciidoc", ".txt" }, ".pdf", "OfficeIMO.AsciiDoc.Pdf", "AsciiDocDocument.Parse(source).Document.ToPdfDocumentResult(options)", "Render bounded AsciiDoc into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("latex-markdown", "LaTeX", "Markdown", OfficeConversionInputKind.Text, new[] { ".tex", ".latex", ".txt" }, ".md", "OfficeIMO.Latex.Markdown", "LatexDocument.Parse(source).Document.ToMarkdownDocumentResult(options)", "Project LaTeX into typed Markdown with diagnostics.", OfficeConversionFidelityKind.Semantic, "LatexToMarkdownResult"),
        Route("markdown-latex", "Markdown", "LaTeX", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".tex", "OfficeIMO.Latex.Markdown", "MarkdownReader.Parse(markdown).ToLatexDocumentResult(options)", "Project typed Markdown into canonical LaTeX with diagnostics.", OfficeConversionFidelityKind.Semantic, "MarkdownToLatexResult"),
        Route("latex-pdf", "LaTeX", "PDF", OfficeConversionInputKind.Text, new[] { ".tex", ".latex", ".txt" }, ".pdf", "OfficeIMO.Latex.Pdf", "LatexDocument.Parse(source).Document.ToPdfDocumentResult(options)", "Render bounded LaTeX into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("onenote-html", "OneNote", "HTML", OfficeConversionInputKind.File, new[] { ".one" }, ".html", "OfficeIMO.OneNote.Html", "section.ToHtmlDocumentResult(projectionOptions, htmlOptions)", "Project a OneNote section into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "HtmlTextConversionResult"),
        Route("html-onenote", "HTML", "OneNote", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".one", "OfficeIMO.OneNote.Html", "HtmlConversionDocument.Parse(html).ToOneNoteSectionResult(options)", "Convert bounded HTML into an editable OneNote section model.", OfficeConversionFidelityKind.Editable, "HtmlToOneNoteSectionResult"),
        Route("onenote-markdown", "OneNote", "Markdown", OfficeConversionInputKind.File, new[] { ".one" }, ".md", "OfficeIMO.OneNote.Markdown", "section.ToMarkdownDocumentResult(options)", "Project a OneNote section into typed Markdown.", OfficeConversionFidelityKind.Semantic, "OneNoteMarkdownConversionResult"),
        Route("onenote-pdf", "OneNote", "PDF", OfficeConversionInputKind.File, new[] { ".one" }, ".pdf", "OfficeIMO.OneNote.Pdf", "section.ToPdfDocumentResult(options)", "Render a OneNote section into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("odt-pdf", "ODT", "PDF", OfficeConversionInputKind.File, new[] { ".odt" }, ".pdf", "OfficeIMO.OpenDocument.Odt.Pdf", "OdtDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render OpenDocument Text into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("ods-pdf", "ODS", "PDF", OfficeConversionInputKind.File, new[] { ".ods" }, ".pdf", "OfficeIMO.OpenDocument.Ods.Pdf", "OdsDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render an OpenDocument spreadsheet into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("odp-pdf", "ODP", "PDF", OfficeConversionInputKind.File, new[] { ".odp" }, ".pdf", "OfficeIMO.OpenDocument.Odp.Pdf", "OdpPresentation.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render an OpenDocument presentation into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("pdf-docx", "PDF", "DOCX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".docx", "OfficeIMO.Word.Pdf", "PdfDocument.Open(stream).ToWordDocumentResult(options)", "Import PDF logical content into an editable Word document with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfWordConversionResult"),
        Route("pdf-xlsx", "PDF", "XLSX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".xlsx", "OfficeIMO.Excel.Pdf", "PdfDocument.Open(stream).ImportTablesToExcelDocumentResult(options)", "Import detected PDF tables into an editable workbook with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfExcelTableImportResult"),
        Route("pdf-pptx", "PDF", "PPTX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".pptx", "OfficeIMO.PowerPoint.Pdf", "PdfDocument.Open(stream).ToPowerPointPresentationResult(options)", "Import PDF pages into an editable presentation profile with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfPowerPointConversionResult"),
        Route("pdf-html", "PDF", "HTML", OfficeConversionInputKind.File, new[] { ".pdf" }, ".html", "OfficeIMO.Html.Pdf", "PdfDocument.Open(stream).ToHtmlResult(options)", "Project PDF logical content into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "PdfHtmlConversionResult"),
        Route("pdf-rtf", "PDF", "RTF", OfficeConversionInputKind.File, new[] { ".pdf" }, ".rtf", "OfficeIMO.Rtf.Pdf", "PdfDocument.Open(stream).ToRtfDocumentResult(options)", "Import PDF logical content into semantic RTF with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfRtfConversionResult"),
        Route("pdf-odt", "PDF", "ODT", OfficeConversionInputKind.File, new[] { ".pdf" }, ".odt", "OfficeIMO.OpenDocument.Odt.Pdf", "PdfDocument.Open(stream).ToOdtDocumentResult(pdfOptions, openDocumentOptions)", "Import PDF logical content into OpenDocument Text with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdtConversionResult"),
        Route("pdf-ods", "PDF", "ODS", OfficeConversionInputKind.File, new[] { ".pdf" }, ".ods", "OfficeIMO.OpenDocument.Ods.Pdf", "PdfDocument.Open(stream).ToOdsDocumentResult(pdfOptions, openDocumentOptions)", "Import detected PDF tables into an OpenDocument spreadsheet with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdsConversionResult"),
        Route("pdf-odp", "PDF", "ODP", OfficeConversionInputKind.File, new[] { ".pdf" }, ".odp", "OfficeIMO.OpenDocument.Odp.Pdf", "PdfDocument.Open(stream).ToOdpPresentationResult(pdfOptions, openDocumentOptions)", "Import PDF pages into an OpenDocument presentation profile with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdpConversionResult"),
        Route("mhtml-pdf", "MHTML", "PDF", OfficeConversionInputKind.File, new[] { ".mhtml", ".mht" }, ".pdf", "OfficeIMO.Html.Pdf", "MhtmlDocument.Load(stream, options).ToPdfDocumentResult(pdfOptions)", "Render a bounded MHTML archive into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("visio-pdf", "Visio", "PDF", OfficeConversionInputKind.File, new[] { ".vsdx" }, ".pdf", "OfficeIMO.Visio.Pdf", "VisioDocument.Load(stream).ToPdfDocumentResult(options)", "Render a Visio drawing into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult")
    };

    private static OfficeConversionCapability Route(
        string id, string source, string target, OfficeConversionInputKind inputKind,
        IEnumerable<string> sourceExtensions, string targetExtension, string packageId,
        string api, string description, OfficeConversionFidelityKind fidelity,
        string resultContract, bool browser = false) =>
        new OfficeConversionCapability(id, source, target, inputKind, sourceExtensions, targetExtension,
            packageId, api, description, fidelity, resultContract, browser);

    private static string NormalizeExtension(string extension) {
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("Extension cannot be empty.", nameof(extension));
        string value = extension.Trim();
        return (value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value).ToLowerInvariant();
    }

    private static string EscapeMarkdown(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");

    private static string EscapeJson(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\").Replace("\"", "\\\"")
        .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
}
