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
    Text,
    /// <summary>The route accepts an authenticated remote document identifier.</summary>
    RemoteResource,
    /// <summary>The route accepts an already materialized public document object.</summary>
    ObjectModel
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

/// <summary>Describes how deeply a conversion route is proven beyond its output model.</summary>
public enum OfficeConversionSupportLevel {
    /// <summary>A bounded subset is implemented and tested; callers must review the documented limitations.</summary>
    Targeted,
    /// <summary>Representative business documents and result artifacts are covered by repeatable tests.</summary>
    Established,
    /// <summary>Complex fixtures, diagnostics, and visual or structural regression evidence cover the route.</summary>
    Advanced,
    /// <summary>Advanced coverage is supplemented by pinned independent-producer or reference-renderer comparisons.</summary>
    ReferenceVerified
}

/// <summary>Describes what a conversion route promises for text and font formatting.</summary>
public enum OfficeConversionTextFormattingKind {
    /// <summary>No route-specific text-formatting promise was supplied.</summary>
    Unspecified,
    /// <summary>The destination retains equivalent editable formatting where its native model can represent it.</summary>
    EditableEquivalent,
    /// <summary>The destination retains equivalent semantic formatting rather than source page geometry.</summary>
    SemanticEquivalent,
    /// <summary>The route preserves only the inline formatting supported by the source and destination syntax profiles.</summary>
    SyntaxSubset,
    /// <summary>The destination retains the rendered appearance through a fixed-layout document.</summary>
    FixedLayoutAppearance,
    /// <summary>The route reconstructs editable or semantic formatting from fixed-layout PDF evidence.</summary>
    ReconstructedFromFixedLayout,
    /// <summary>The destination retains rendered vector text and decoration attributes, not editable source semantics.</summary>
    VectorAppearance,
    /// <summary>The destination retains rendered pixels and cannot retain editable text or font semantics.</summary>
    FlattenedRaster,
    /// <summary>The format carries tabular values only and has no font or text-style model.</summary>
    DataOnly
}

/// <summary>One package-neutral conversion route exposed by OfficeIMO.</summary>
public sealed class OfficeConversionCapability {
    /// <summary>Creates a conversion capability using conservative support defaults.</summary>
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
        bool agentDiscoverable = true)
        : this(
            id,
            source,
            target,
            inputKind,
            sourceExtensions,
            targetExtension,
            packageId,
            api,
            description,
            fidelity,
            resultContract,
            browserAvailable,
            agentDiscoverable,
            OfficeConversionSupportLevel.Targeted,
            "No route-specific evidence summary was supplied.",
            "Review the conversion report and package documentation before relying on unsupported constructs.") {
    }

    /// <summary>Creates a conversion capability with an explicit evidence-based support assessment.</summary>
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
        bool browserAvailable,
        bool agentDiscoverable,
        OfficeConversionSupportLevel supportLevel,
        string supportEvidence,
        string knownLimitations)
        : this(
            id,
            source,
            target,
            inputKind,
            sourceExtensions,
            targetExtension,
            packageId,
            api,
            description,
            fidelity,
            resultContract,
            browserAvailable,
            agentDiscoverable,
            supportLevel,
            supportEvidence,
            knownLimitations,
            OfficeConversionTextFormattingKind.Unspecified,
            null) {
    }

    /// <summary>Creates a conversion capability with an explicit evidence-based support and text-formatting assessment.</summary>
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
        bool browserAvailable,
        bool agentDiscoverable,
        OfficeConversionSupportLevel supportLevel,
        string supportEvidence,
        string knownLimitations,
        OfficeConversionTextFormattingKind textFormatting = OfficeConversionTextFormattingKind.Unspecified,
        string? textFormattingContract = null) {
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
        SupportLevel = supportLevel;
        SupportEvidence = string.IsNullOrWhiteSpace(supportEvidence)
            ? throw new ArgumentException("Support evidence cannot be empty.", nameof(supportEvidence))
            : supportEvidence.Trim();
        KnownLimitations = string.IsNullOrWhiteSpace(knownLimitations)
            ? throw new ArgumentException("Known limitations cannot be empty.", nameof(knownLimitations))
            : knownLimitations.Trim();
        TextFormatting = textFormatting;
        TextFormattingContract = string.IsNullOrWhiteSpace(textFormattingContract)
            ? "Preserves only text formatting representable by the destination; consult conversion diagnostics for approximations and omissions."
            : textFormattingContract!.Trim();
    }

    /// <summary>Gets the stable route identifier.</summary>
    public string Id { get; }
    /// <summary>Gets the source format label.</summary>
    public string Source { get; }
    /// <summary>Gets the destination format label.</summary>
    public string Target { get; }
    /// <summary>Gets how the route accepts source content.</summary>
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
    /// <summary>Gets the depth of repeatable evidence behind the route.</summary>
    public OfficeConversionSupportLevel SupportLevel { get; }
    /// <summary>Gets a concise summary of the evidence supporting the assigned level.</summary>
    public string SupportEvidence { get; }
    /// <summary>Gets the important unsupported or intentionally simplified scope.</summary>
    public string KnownLimitations { get; }
    /// <summary>Gets the route's text and font formatting fidelity classification.</summary>
    public OfficeConversionTextFormattingKind TextFormatting { get; }
    /// <summary>Gets the explicit text and font formatting promise for the route.</summary>
    public string TextFormattingContract { get; }

    private static string NormalizeExtension(string extension) {
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("Extensions cannot be empty.", nameof(extension));
        string value = extension.Trim();
        return (value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value).ToLowerInvariant();
    }
}

/// <summary>The shared OfficeIMO conversion route catalog used by packages, agents, and browser surfaces.</summary>
public static class OfficeConversionCapabilityCatalog {
    /// <summary>Gets the capability schema version.</summary>
    public const int SchemaVersion = 7;

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
        markdown.AppendLine("Use this table to find the focused package, output model, proven support level, known limits, and result type for a source-to-target conversion.");
        markdown.AppendLine();
        markdown.Append("Schema version: ").Append(SchemaVersion).AppendLine();
        markdown.AppendLine();
        markdown.AppendLine("| Route | Source | Target | Package | Output model | Text formatting | Typography contract | Support | Evidence | Known limits | Browser | API | Result type | What it does |");
        markdown.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        foreach (OfficeConversionCapability route in All) {
            markdown.Append("| ").Append(EscapeMarkdown(route.Id))
                .Append(" | ").Append(EscapeMarkdown(route.Source))
                .Append(" | ").Append(EscapeMarkdown(route.Target))
                .Append(" | ").Append(EscapeMarkdown(route.PackageId))
                .Append(" | ").Append(route.Fidelity)
                .Append(" | ").Append(route.TextFormatting)
                .Append(" | ").Append(EscapeMarkdown(route.TextFormattingContract))
                .Append(" | ").Append(route.SupportLevel)
                .Append(" | ").Append(EscapeMarkdown(route.SupportEvidence))
                .Append(" | ").Append(EscapeMarkdown(route.KnownLimitations))
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
                .Append(i3).Append("\"textFormatting\":\"").Append(route.TextFormatting).Append("\",").Append(newline)
                .Append(i3).Append("\"textFormattingContract\":\"").Append(EscapeJson(route.TextFormattingContract)).Append("\",").Append(newline)
                .Append(i3).Append("\"supportLevel\":\"").Append(route.SupportLevel).Append("\",").Append(newline)
                .Append(i3).Append("\"supportEvidence\":\"").Append(EscapeJson(route.SupportEvidence)).Append("\",").Append(newline)
                .Append(i3).Append("\"knownLimitations\":\"").Append(EscapeJson(route.KnownLimitations)).Append("\",").Append(newline)
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

    private static OfficeConversionCapability[] CreateRoutes() {
        var routes = new List<OfficeConversionCapability> {
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
        Route("rtf-docx", "RTF", "DOCX", OfficeConversionInputKind.File, new[] { ".rtf" }, ".docx", "OfficeIMO.Word.Rtf", "RtfDocument.LoadResult(stream, readOptions).ToWordDocumentResult(sourcePath)", "Convert bounded RTF into an editable Word document while retaining read and conversion diagnostics.", OfficeConversionFidelityKind.Editable, "RtfConversionResult<WordDocument>"),
        Route("xlsx-html", "XLSX", "HTML", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".html", "OfficeIMO.Excel.Html", "ExcelDocument.Load(stream).ToHtmlResult(options)", "Project workbook content into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "HtmlTextConversionResult"),
        Route("xlsx-ods", "XLSX", "ODS", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".ods", "OfficeIMO.Excel.OpenDocument", "ExcelDocument.Load(stream).ToOpenDocumentResult(options)", "Convert an editable workbook to OpenDocument Spreadsheet.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<OdsDocument>"),
        Route("ods-xlsx", "ODS", "XLSX", OfficeConversionInputKind.File, new[] { ".ods" }, ".xlsx", "OfficeIMO.Excel.OpenDocument", "OdsDocument.Load(stream).ToExcelDocumentResult(options)", "Convert OpenDocument Spreadsheet into an editable workbook.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<ExcelDocument>"),
        Route("pptx-html", "PPTX", "HTML", OfficeConversionInputKind.File, new[] { ".pptx" }, ".html", "OfficeIMO.PowerPoint.Html", "PowerPointPresentation.Load(stream).ToHtmlResult(options)", "Project presentation slides into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "PowerPointToHtmlResult"),
        Route("pptx-odp", "PPTX", "ODP", OfficeConversionInputKind.File, new[] { ".pptx" }, ".odp", "OfficeIMO.PowerPoint.OpenDocument", "PowerPointPresentation.Load(stream).ToOpenDocumentResult(options)", "Convert an editable presentation to OpenDocument Presentation.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<OdpPresentation>"),
        Route("odp-pptx", "ODP", "PPTX", OfficeConversionInputKind.File, new[] { ".odp" }, ".pptx", "OfficeIMO.PowerPoint.OpenDocument", "OdpPresentation.Load(stream).ToPowerPointPresentationResult(options)", "Convert OpenDocument Presentation into an editable presentation.", OfficeConversionFidelityKind.Editable, "OdfConversionResult<PowerPointPresentation>"),
        Route("markdown-pdf", "Markdown", "PDF", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".pdf", "OfficeIMO.Markdown.Pdf", "MarkdownDoc.Parse(markdown).ToPdfDocumentResult(options)", "Render Markdown into PDF with structured conversion warnings.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("rtf-markdown", "RTF", "Markdown", OfficeConversionInputKind.File, new[] { ".rtf" }, ".md", "OfficeIMO.Rtf.Markdown", "RtfDocument.LoadResult(stream, readOptions).Document.ToMarkdownResult(options)", "Project bounded RTF into portable Markdown.", OfficeConversionFidelityKind.Semantic, "RtfConversionResult<string>"),
        Route("rtf-pdf", "RTF", "PDF", OfficeConversionInputKind.File, new[] { ".rtf" }, ".pdf", "OfficeIMO.Rtf.Pdf", "RtfDocument.LoadResult(stream, readOptions).Document.ToPdfDocumentResult(options)", "Render bounded RTF into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("markdown-rtf", "Markdown", "RTF", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".rtf", "OfficeIMO.Rtf.Markdown", "MarkdownDoc.Parse(markdown).ToRtfDocumentResult(options)", "Convert typed Markdown into semantic RTF with loss diagnostics.", OfficeConversionFidelityKind.Editable, "RtfConversionResult<RtfDocument>"),
        Route("html-docx", "HTML", "DOCX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".docx", "OfficeIMO.Word.Html", "HtmlConversionDocument.Parse(html).ToWordDocumentResult(options)", "Convert bounded HTML into an editable Word document.", OfficeConversionFidelityKind.Editable, "HtmlToWordResult"),
        Route("html-xlsx", "HTML", "XLSX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".xlsx", "OfficeIMO.Excel.Html", "HtmlConversionDocument.Parse(html).ToExcelDocumentResult(options)", "Convert bounded HTML tables and content into an editable workbook.", OfficeConversionFidelityKind.Editable, "HtmlToExcelResult"),
        Route("html-pptx", "HTML", "PPTX", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".pptx", "OfficeIMO.PowerPoint.Html", "HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(options)", "Convert bounded HTML into an editable presentation.", OfficeConversionFidelityKind.Editable, "HtmlToPowerPointResult"),
        Route("html-rtf", "HTML", "RTF", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".rtf", "OfficeIMO.Html.Rtf", "HtmlConversionDocument.Parse(html).ToRtfDocumentResult(options)", "Convert bounded HTML into semantic RTF.", OfficeConversionFidelityKind.Editable, "HtmlToRtfResult"),
        Route("rtf-html", "RTF", "HTML", OfficeConversionInputKind.File, new[] { ".rtf" }, ".html", "OfficeIMO.Html.Rtf", "RtfDocument.LoadResult(stream, readOptions).Document.ToHtmlResult(options)", "Render bounded RTF through an explicit safe HTML profile.", OfficeConversionFidelityKind.Semantic, "RtfToHtmlResult"),
        Route("asciidoc-markdown", "AsciiDoc", "Markdown", OfficeConversionInputKind.Text, new[] { ".adoc", ".asciidoc", ".txt" }, ".md", "OfficeIMO.AsciiDoc.Markdown", "AsciiDocDocument.ParseResult(source).Document.ToMarkdownDocumentResult(options)", "Project AsciiDoc into typed Markdown with diagnostics.", OfficeConversionFidelityKind.Semantic, "AsciiDocToMarkdownResult"),
        Route("markdown-asciidoc", "Markdown", "AsciiDoc", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".adoc", "OfficeIMO.AsciiDoc.Markdown", "MarkdownDoc.Parse(markdown).ToAsciiDocDocumentResult(options)", "Project typed Markdown into canonical AsciiDoc with diagnostics.", OfficeConversionFidelityKind.Semantic, "MarkdownToAsciiDocResult"),
        Route("asciidoc-pdf", "AsciiDoc", "PDF", OfficeConversionInputKind.Text, new[] { ".adoc", ".asciidoc", ".txt" }, ".pdf", "OfficeIMO.AsciiDoc.Pdf", "AsciiDocDocument.ParseResult(source).Document.ToPdfDocumentResult(options)", "Render bounded AsciiDoc into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("latex-markdown", "LaTeX", "Markdown", OfficeConversionInputKind.Text, new[] { ".tex", ".latex", ".txt" }, ".md", "OfficeIMO.Latex.Markdown", "LatexDocument.ParseResult(source).Document.ToMarkdownDocumentResult(options)", "Project LaTeX into typed Markdown with diagnostics.", OfficeConversionFidelityKind.Semantic, "LatexToMarkdownResult"),
        Route("markdown-latex", "Markdown", "LaTeX", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".tex", "OfficeIMO.Latex.Markdown", "MarkdownDoc.Parse(markdown).ToLatexDocumentResult(options)", "Project typed Markdown into canonical LaTeX with diagnostics.", OfficeConversionFidelityKind.Semantic, "MarkdownToLatexResult"),
        Route("latex-pdf", "LaTeX", "PDF", OfficeConversionInputKind.Text, new[] { ".tex", ".latex", ".txt" }, ".pdf", "OfficeIMO.Latex.Pdf", "LatexDocument.ParseResult(source).Document.ToPdfDocumentResult(options)", "Render bounded LaTeX into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("onenote-html", "OneNote", "HTML", OfficeConversionInputKind.File, new[] { ".one" }, ".html", "OfficeIMO.OneNote.Html", "section.ToHtmlDocumentResult(projectionOptions, htmlOptions)", "Project a OneNote section into reviewable HTML.", OfficeConversionFidelityKind.Semantic, "HtmlTextConversionResult"),
        Route("html-onenote", "HTML", "OneNote", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".one", "OfficeIMO.OneNote.Html", "HtmlConversionDocument.Parse(html).ToOneNoteSectionResult(options)", "Convert bounded HTML into an editable OneNote section model.", OfficeConversionFidelityKind.Editable, "HtmlToOneNoteSectionResult"),
        Route("onenote-markdown", "OneNote", "Markdown", OfficeConversionInputKind.File, new[] { ".one" }, ".md", "OfficeIMO.OneNote.Markdown", "section.ToMarkdownDocumentResult(options)", "Project a OneNote section into typed Markdown.", OfficeConversionFidelityKind.Semantic, "OneNoteMarkdownConversionResult"),
        Route("onenote-pdf", "OneNote", "PDF", OfficeConversionInputKind.File, new[] { ".one" }, ".pdf", "OfficeIMO.OneNote.Pdf", "section.ToPdfDocumentResult(options)", "Render a OneNote section into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("odt-pdf", "ODT", "PDF", OfficeConversionInputKind.File, new[] { ".odt" }, ".pdf", "OfficeIMO.OpenDocument.Odt.Pdf", "OdtDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render OpenDocument Text into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("ods-pdf", "ODS", "PDF", OfficeConversionInputKind.File, new[] { ".ods" }, ".pdf", "OfficeIMO.OpenDocument.Ods.Pdf", "OdsDocument.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render an OpenDocument spreadsheet into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("odp-pdf", "ODP", "PDF", OfficeConversionInputKind.File, new[] { ".odp" }, ".pdf", "OfficeIMO.OpenDocument.Odp.Pdf", "OdpPresentation.Load(stream).ToPdfDocumentResult(conversionOptions, pdfOptions)", "Render an OpenDocument presentation into a fixed-layout PDF with source diagnostics.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("pdf-docx", "PDF", "DOCX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".docx", "OfficeIMO.Word.Pdf", "PdfDocument.Load(stream).ToWordDocumentResult(options)", "Import PDF logical content into an editable Word document with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfWordConversionResult", browser: true),
        Route("pdf-xlsx", "PDF", "XLSX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".xlsx", "OfficeIMO.Excel.Pdf", "PdfDocument.Load(stream).ImportTablesToExcelDocumentResult(options)", "Import detected PDF tables into an editable workbook with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfExcelTableImportResult", browser: true),
        Route("pdf-pptx", "PDF", "PPTX", OfficeConversionInputKind.File, new[] { ".pdf" }, ".pptx", "OfficeIMO.PowerPoint.Pdf", "PdfDocument.Load(stream).ToPowerPointPresentationResult(options)", "Reconstruct supported PDF content as native slide objects, or select explicit visual, hybrid, and tables-only projections with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfPowerPointConversionResult", browser: true),
        Route("pdf-html", "PDF", "HTML", OfficeConversionInputKind.File, new[] { ".pdf" }, ".html", "OfficeIMO.Html.Pdf", "PdfDocument.Load(stream).ToHtmlResult(options)", "Project PDF content into semantic or positioned-review HTML with explicit diagnostics.", OfficeConversionFidelityKind.Semantic, "PdfHtmlConversionResult", browser: true),
        Route("pdf-png", "PDF", "PNG", OfficeConversionInputKind.File, new[] { ".pdf" }, ".png", "OfficeIMO.Pdf", "PdfDocument.Load(stream).Render.ExportImages(OfficeImageExportFormat.Png, options)", "Render every PDF page as a detailed PNG with page-level diagnostics; multi-page results are packaged as a ZIP.", OfficeConversionFidelityKind.FixedLayout, "IReadOnlyList<OfficeImageExportResult>", browser: true),
        Route("pdf-rtf", "PDF", "RTF", OfficeConversionInputKind.File, new[] { ".pdf" }, ".rtf", "OfficeIMO.Rtf.Pdf", "PdfDocument.Load(stream).ToRtfDocumentResult(options)", "Import PDF logical content into semantic RTF with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfRtfConversionResult"),
        Route("pdf-odt", "PDF", "ODT", OfficeConversionInputKind.File, new[] { ".pdf" }, ".odt", "OfficeIMO.OpenDocument.Odt.Pdf", "PdfDocument.Load(stream).ToOdtDocumentResult(pdfOptions, openDocumentOptions)", "Import PDF logical content into OpenDocument Text with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdtConversionResult"),
        Route("pdf-ods", "PDF", "ODS", OfficeConversionInputKind.File, new[] { ".pdf" }, ".ods", "OfficeIMO.OpenDocument.Ods.Pdf", "PdfDocument.Load(stream).ToOdsDocumentResult(pdfOptions, openDocumentOptions)", "Import detected PDF tables into an OpenDocument spreadsheet with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdsConversionResult"),
        Route("pdf-odp", "PDF", "ODP", OfficeConversionInputKind.File, new[] { ".pdf" }, ".odp", "OfficeIMO.OpenDocument.Odp.Pdf", "PdfDocument.Load(stream).ToOdpPresentationResult(pdfOptions, openDocumentOptions)", "Import PDF pages into an OpenDocument presentation profile with diagnostics.", OfficeConversionFidelityKind.Editable, "PdfOdpConversionResult"),
        Route("mhtml-pdf", "MHTML", "PDF", OfficeConversionInputKind.File, new[] { ".mhtml", ".mht" }, ".pdf", "OfficeIMO.Mhtml.Pdf", "MhtmlDocument.Load(stream, options).ToPdfDocumentResult(pdfOptions)", "Render a bounded MHTML archive into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult"),
        Route("docx-google-docs", "DOCX", "Google Docs", OfficeConversionInputKind.File, new[] { ".docx" }, ".gdoc", "OfficeIMO.Word.GoogleDocs", "WordDocument.Load(stream).ExportToGoogleDocsAsync(session, options)", "Export editable Word content to an authenticated Google document.", OfficeConversionFidelityKind.Editable, "Task<GoogleDocumentReference>"),
        Route("google-docs-docx", "Google Docs", "DOCX", OfficeConversionInputKind.RemoteResource, new[] { ".gdoc" }, ".docx", "OfficeIMO.Word.GoogleDocs", "session.ImportGoogleDocAsync(documentId, options)", "Import an authenticated Google document through native projection or Drive DOCX conversion.", OfficeConversionFidelityKind.Editable, "Task<GoogleDocsImportResult>"),
        Route("xlsx-google-sheets", "XLSX", "Google Sheets", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".gsheet", "OfficeIMO.Excel.GoogleSheets", "ExcelDocument.Load(stream).ExportToGoogleSheetsAsync(session, options)", "Export an editable workbook to an authenticated Google spreadsheet.", OfficeConversionFidelityKind.Editable, "Task<GoogleSpreadsheetReference>"),
        Route("google-sheets-xlsx", "Google Sheets", "XLSX", OfficeConversionInputKind.RemoteResource, new[] { ".gsheet" }, ".xlsx", "OfficeIMO.Excel.GoogleSheets", "session.ImportGoogleSheetAsync(spreadsheetId, options)", "Import an authenticated Google spreadsheet through native projection or Drive XLSX conversion.", OfficeConversionFidelityKind.Editable, "Task<GoogleSheetsImportResult>"),
        Route("pptx-google-slides", "PPTX", "Google Slides", OfficeConversionInputKind.File, new[] { ".pptx" }, ".gslides", "OfficeIMO.PowerPoint.GoogleSlides", "PowerPointPresentation.Load(stream).ExportToGoogleSlidesAsync(session, options)", "Export an editable presentation to an authenticated Google presentation.", OfficeConversionFidelityKind.Editable, "Task<GooglePresentationReference>"),
        Route("google-slides-pptx", "Google Slides", "PPTX", OfficeConversionInputKind.RemoteResource, new[] { ".gslides" }, ".pptx", "OfficeIMO.PowerPoint.GoogleSlides", "session.ImportGoogleSlidesAsync(presentationId, options)", "Import an authenticated Google presentation through native projection or Drive PPTX conversion.", OfficeConversionFidelityKind.Editable, "Task<GoogleSlidesImportResult>"),
        Route("adf-markdown", "ADF", "Markdown", OfficeConversionInputKind.Text, new[] { ".adf", ".json" }, ".md", "OfficeIMO.Adf", "AdfConverter.ToMarkdown(AdfDocument.Parse(json), options)", "Project Atlassian Document Format JSON into portable Markdown with fidelity diagnostics.", OfficeConversionFidelityKind.Semantic, "AdfConversionResult<string>"),
        Route("markdown-adf", "Markdown", "ADF", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".adf", "OfficeIMO.Adf", "AdfConverter.FromMarkdown(markdown)", "Convert typed Markdown into Atlassian Document Format JSON.", OfficeConversionFidelityKind.Semantic, "AdfConversionResult<AdfDocument>"),
        Route("adf-html", "ADF", "HTML", OfficeConversionInputKind.Text, new[] { ".adf", ".json" }, ".html", "OfficeIMO.Adf", "AdfConverter.ToHtml(AdfDocument.Parse(json), htmlOptions, options)", "Render Atlassian Document Format through the canonical Markdown and HTML models.", OfficeConversionFidelityKind.Semantic, "AdfConversionResult<string>"),
        Route("html-adf", "HTML", "ADF", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".adf", "OfficeIMO.Adf", "AdfConverter.FromHtml(html, options)", "Project bounded HTML through Markdown into Atlassian Document Format.", OfficeConversionFidelityKind.Semantic, "AdfConversionResult<AdfDocument>"),
        Route("markdown-confluence", "Markdown", "Confluence", OfficeConversionInputKind.Text, new[] { ".md", ".markdown", ".txt" }, ".adf", "OfficeIMO.Confluence", "ConfluenceContentConverter.FromMarkdown(markdown, format)", "Create a Confluence ADF or storage body from Markdown.", OfficeConversionFidelityKind.Semantic, "ConfluenceContentConversionResult<ConfluencePageBody>"),
        Route("html-confluence", "HTML", "Confluence", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, ".adf", "OfficeIMO.Confluence", "ConfluenceContentConverter.FromHtml(html, format)", "Create a Confluence storage or ADF body from bounded HTML.", OfficeConversionFidelityKind.Semantic, "ConfluenceContentConversionResult<ConfluencePageBody>"),
        Route("confluence-markdown", "Confluence", "Markdown", OfficeConversionInputKind.ObjectModel, new[] { ".adf", ".json" }, ".md", "OfficeIMO.Confluence", "ConfluenceContentConverter.ToMarkdown(page)", "Project a materialized Confluence page body to Markdown with fidelity diagnostics.", OfficeConversionFidelityKind.Semantic, "ConfluenceContentConversionResult<string>"),
        Route("confluence-html", "Confluence", "HTML", OfficeConversionInputKind.ObjectModel, new[] { ".adf", ".json" }, ".html", "OfficeIMO.Confluence", "ConfluenceContentConverter.ToHtml(page)", "Project a materialized Confluence page body to HTML with fidelity diagnostics.", OfficeConversionFidelityKind.Semantic, "ConfluenceContentConversionResult<string>"),
        Route("csv-xlsx", "CSV", "XLSX", OfficeConversionInputKind.File, new[] { ".csv", ".tsv" }, ".xlsx", "OfficeIMO.Excel.Csv", "CsvDocument.Load(stream).ToExcelDocument(options)", "Import delimited values into an editable workbook.", OfficeConversionFidelityKind.Editable, "ExcelDocument"),
        Route("xlsx-csv", "XLSX", "CSV", OfficeConversionInputKind.File, new[] { ".xlsx" }, ".csv", "OfficeIMO.Excel.Csv", "ExcelDocument.Load(stream).Sheets[0].ToCsv(options)", "Export a worksheet used range as delimited values.", OfficeConversionFidelityKind.Semantic, "string"),
        Route("officemarkup-docx", "OfficeIMO Markup", "DOCX", OfficeConversionInputKind.Text, new[] { ".omd", ".office.md" }, ".docx", "OfficeIMO.Markup.Word", "OfficeMarkupParser.Parse(markup, options).Document.ToWordDocumentResult(exportOptions)", "Render document-profile OfficeIMO Markup into an editable Word document.", OfficeConversionFidelityKind.Editable, "OfficeMarkupConversionResult<WordDocument>"),
        Route("officemarkup-xlsx", "OfficeIMO Markup", "XLSX", OfficeConversionInputKind.Text, new[] { ".omd", ".office.md" }, ".xlsx", "OfficeIMO.Markup.Excel", "OfficeMarkupParser.Parse(markup, options).Document.ToExcelDocumentResult(exportOptions)", "Render workbook-profile OfficeIMO Markup into an editable Excel workbook.", OfficeConversionFidelityKind.Editable, "OfficeMarkupConversionResult<ExcelDocument>"),
        Route("officemarkup-pptx", "OfficeIMO Markup", "PPTX", OfficeConversionInputKind.Text, new[] { ".omd", ".office.md" }, ".pptx", "OfficeIMO.Markup.PowerPoint", "OfficeMarkupParser.Parse(markup, options).Document.ToPowerPointPresentationResult(exportOptions)", "Render presentation-profile OfficeIMO Markup into an editable PowerPoint presentation.", OfficeConversionFidelityKind.Editable, "OfficeMarkupPowerPointConversionResult"),
        Route("visio-pdf", "Visio", "PDF", OfficeConversionInputKind.File, new[] { ".vsdx" }, ".pdf", "OfficeIMO.Visio.Pdf", "VisioDocument.Load(stream).ToPdfDocumentResult(options)", "Render a Visio drawing into a fixed-layout PDF.", OfficeConversionFidelityKind.FixedLayout, "PdfDocumentConversionResult")
        };

        AddImageRoutes(routes, "docx", "DOCX", OfficeConversionInputKind.File, new[] { ".docx" }, "OfficeIMO.Word", "WordDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "xlsx", "XLSX", OfficeConversionInputKind.File, new[] { ".xlsx" }, "OfficeIMO.Excel", "ExcelDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "pptx", "PPTX", OfficeConversionInputKind.File, new[] { ".pptx" }, "OfficeIMO.PowerPoint", "PowerPointPresentation.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "html", "HTML", OfficeConversionInputKind.Text, new[] { ".html", ".htm", ".txt" }, "OfficeIMO.Html", "HtmlConversionDocument.Parse(html).ExportImages(format, options)");
        AddImageRoutes(routes, "onenote", "OneNote", OfficeConversionInputKind.File, new[] { ".one" }, "OfficeIMO.OneNote", "OneNoteSectionReader.Read(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "visio", "Visio", OfficeConversionInputKind.File, new[] { ".vsdx" }, "OfficeIMO.Visio", "VisioDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "email", "Email", OfficeConversionInputKind.File, new[] { ".eml" }, "OfficeIMO.Email.Image", "EmailDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "epub", "EPUB", OfficeConversionInputKind.File, new[] { ".epub" }, "OfficeIMO.Epub.Image", "EpubDocument.Load(stream, readOptions).ExportImages(format, options)");
        AddImageRoutes(routes, "odt", "ODT", OfficeConversionInputKind.File, new[] { ".odt" }, "OfficeIMO.Word.OpenDocument", "OdtDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "ods", "ODS", OfficeConversionInputKind.File, new[] { ".ods" }, "OfficeIMO.Excel.OpenDocument", "OdsDocument.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "odp", "ODP", OfficeConversionInputKind.File, new[] { ".odp" }, "OfficeIMO.PowerPoint.OpenDocument", "OdpPresentation.Load(stream).ExportImages(format, options)");
        AddImageRoutes(routes, "pdf", "PDF", OfficeConversionInputKind.File, new[] { ".pdf" }, "OfficeIMO.Pdf", "PdfDocument.Load(stream).Render.ExportImages(format, options)");
        return routes.ToArray();
    }

    private static void AddImageRoutes(
        ICollection<OfficeConversionCapability> routes,
        string idPrefix,
        string source,
        OfficeConversionInputKind inputKind,
        IEnumerable<string> sourceExtensions,
        string packageId,
        string api) {
        var formats = new[] {
            (Id: "png", Label: "PNG", Extension: ".png"),
            (Id: "svg", Label: "SVG", Extension: ".svg"),
            (Id: "jpeg", Label: "JPEG", Extension: ".jpg"),
            (Id: "tiff", Label: "TIFF", Extension: ".tiff"),
            (Id: "webp", Label: "WebP", Extension: ".webp")
        };
        foreach (var format in formats) {
            string id = idPrefix + "-" + format.Id;
            if (routes.Any(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase))) continue;
            routes.Add(Route(
                id,
                source,
                format.Label,
                inputKind,
                sourceExtensions,
                format.Extension,
                packageId,
                api,
                "Render " + source + " content as " + format.Label + " images through the shared image-export contract.",
                OfficeConversionFidelityKind.FixedLayout,
                "IReadOnlyList<OfficeImageExportResult>"));
        }
    }

    private static OfficeConversionCapability Route(
        string id, string source, string target, OfficeConversionInputKind inputKind,
        IEnumerable<string> sourceExtensions, string targetExtension, string packageId,
        string api, string description, OfficeConversionFidelityKind fidelity,
        string resultContract, bool browser = false) =>
        CreateRoute(id, source, target, inputKind, sourceExtensions, targetExtension,
            packageId, api, description, fidelity, resultContract, browser);

    private static OfficeConversionCapability CreateRoute(
        string id, string source, string target, OfficeConversionInputKind inputKind,
        IEnumerable<string> sourceExtensions, string targetExtension, string packageId,
        string api, string description, OfficeConversionFidelityKind fidelity,
        string resultContract, bool browser) {
        OfficeConversionSupportAssessment support = OfficeConversionSupportAssessments.Get(id);
        (OfficeConversionTextFormattingKind textFormatting, string textFormattingContract) = GetTextFormattingContract(source, target);
        return new OfficeConversionCapability(
            id, source, target, inputKind, sourceExtensions, targetExtension,
            packageId, api, description, fidelity, resultContract, browser,
            agentDiscoverable: true,
            supportLevel: support.Level,
            supportEvidence: support.Evidence,
            knownLimitations: support.KnownLimitations,
            textFormatting: textFormatting,
            textFormattingContract: textFormattingContract);
    }

    private static (OfficeConversionTextFormattingKind Kind, string Contract) GetTextFormattingContract(
        string source,
        string target) {
        if (string.Equals(target, "SVG", StringComparison.Ordinal)) {
            return (OfficeConversionTextFormattingKind.VectorAppearance,
                "Preserves rendered font, color, weight, italic, decoration, and script appearance as SVG text and vector graphics; source-native editable semantics are not retained.");
        }
        if (target is "PNG" or "JPEG" or "TIFF" or "WebP") {
            return (OfficeConversionTextFormattingKind.FlattenedRaster,
                "Preserves rendered font styling as pixels; text, casing metadata, decoration variants, and script semantics are no longer editable.");
        }
        if (string.Equals(source, "PDF", StringComparison.Ordinal)) {
            return (OfficeConversionTextFormattingKind.ReconstructedFromFixedLayout,
                "Reconstructs supported text and formatting from PDF logical or positioned content; it cannot recover source-only font semantics absent from the PDF.");
        }
        if (string.Equals(target, "PDF", StringComparison.Ordinal)) {
            return (OfficeConversionTextFormattingKind.FixedLayoutAppearance,
                "Preserves supported font styling as fixed-layout PDF text and graphics; conversion diagnostics identify source semantics the managed renderer approximates or omits.");
        }
        if (source is "Markdown" or "AsciiDoc" or "LaTeX" || target is "Markdown" or "AsciiDoc" or "LaTeX") {
            return (OfficeConversionTextFormattingKind.SyntaxSubset,
                "Preserves only emphasis, strike, script, and inline styling represented by the supported source and destination syntax profiles; arbitrary family, size, color, casing metadata, and underline variants are not portable.");
        }
        if (source == "CSV" || target == "CSV") {
            return (OfficeConversionTextFormattingKind.DataOnly,
                "CSV and TSV carry values, delimiters, and records only; font family, size, color, emphasis, decoration, scripts, casing metadata, and layout are intentionally not representable.");
        }
        if (source == "OfficeIMO Markup") {
            return (OfficeConversionTextFormattingKind.EditableEquivalent,
                "Authors editable native typography, including family, size, color, emphasis, decoration, script, and casing, in the generated Office document; diagnostics identify profile-specific approximations and omissions.");
        }
        if (source is "ADF" or "Confluence" || target is "ADF" or "Confluence") {
            return (OfficeConversionTextFormattingKind.SyntaxSubset,
                "Preserves the text styling represented by the source and destination profiles; unsupported native decoration variants, arbitrary CSS presentation, and format-specific layout are reported or simplified.");
        }
        if (target == "HTML") {
            return (OfficeConversionTextFormattingKind.SemanticEquivalent,
                "Preserves representable font family, size, color, weight, italic, decoration, script, and casing semantics in HTML/CSS, with OfficeIMO metadata for richer native variants where supported.");
        }
        return (OfficeConversionTextFormattingKind.EditableEquivalent,
            "Preserves equivalent native font family, size, color, weight, italic, decoration, script, and casing semantics where the destination supports them; diagnostics identify approximations and omissions.");
    }

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
