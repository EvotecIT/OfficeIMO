using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Mcp;

internal sealed partial class OfficeImoMcpTools {
    [McpServerTool(Name = "officeimo_pdf_compare", Title = "Export ordinal PDF comparison", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Compare independently selected PDF page sequences in caller order and publish a bounded standalone HTML gallery. Requires both source ids issued by inspect/search in this process, an explicit separate output and explicit overwrite permission. Sources remain unchanged. Reports rendered appearance only, with original page numbers and unmatched selected pages; no semantic or moved-page detection. Returns bounded artifact metadata, never embedded images or document text.")]
    public Task<CallToolResult> ComparePdfAsync(
        [Description("Expected PDF source id returned by inspect/search.")] string sourceId,
        [Description("Actual PDF source id returned by inspect/search.")] string comparisonSourceId,
        [Description("Separate destination .html within allowed roots.")] string outputPath,
        [Description("Expected ordered selection, e.g. 20-24,last. Null selects all pages. At most 100 selected pages per side.")] string? expectedPages = null,
        [Description("Actual ordered selection, e.g. 21-25,last. Null selects all pages. Selection positions are paired.")] string? actualPages = null,
        [Description("Host-admitted expected-PDF password environment variable.")] string? passwordEnvironmentVariable = null,
        [Description("Host-admitted actual-PDF password environment variable. Null reuses expected password.")] string? comparisonPasswordEnvironmentVariable = null,
        [Description("Explicit permission to replace an existing report.")] bool overwrite = false,
        [Description("Maximum input bytes per document, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum output bytes, 1-536870912.")] long maximumOutputBytes = 100663296,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecuteSafePdfAsync(() => _service.ComparePdfAsync(sourceId, comparisonSourceId,
            outputPath, Settings(expectedPages, passwordEnvironmentVariable, overwrite, 100, 150, maximumInputBytes, maximumOutputBytes),
            actualPages, comparisonPasswordEnvironmentVariable, maxOutputCharacters, cancellationToken),
            result => result.Summary ?? "PDF comparison: " + result.Status + "; " + result.ArtifactCount + " artifact(s).", result => result.Succeeded);

    [McpServerTool(Name = "officeimo_pdf_redline", Title = "Export a PDF text redline", ReadOnly = false,
        Destructive = true, Idempotent = true, OpenWorld = false, UseStructuredContent = true, OutputSchemaType = typeof(AgentPdfWorkflowResult))]
    [Description("Compare the text of two PDFs word by word in reading order and write a separate copy of the newer PDF with review annotations: highlights on inserted, replaced and moved text, and carets where text was deleted, with notes carrying the old text. Optionally writes the full change list as JSON (page numbers and page rectangles) and a side-by-side HTML report. Requires both source ids issued by inspect/search in this process, explicit separate outputs and explicit overwrite permission. Sources remain unchanged. PDFs without a text layer are refused; run OCR first. Returns change counts and artifact metadata, never document text.")]
    public Task<CallToolResult> RedlinePdfAsync(
        [Description("Older PDF source id returned by inspect/search.")] string sourceId,
        [Description("Newer PDF source id returned by inspect/search; its copy receives the annotations.")] string comparisonSourceId,
        [Description("Separate destination .pdf within allowed roots for the annotated copy.")] string outputPath,
        [Description("Optional separate destination .json for the complete change list.")] string? changesPath = null,
        [Description("Optional separate destination .html for the side-by-side report.")] string? reportPath = null,
        [Description("Word alignment: myers (minimal, default), patience or histogram.")] string algorithm = "myers",
        [Description("Ignore letter case.")] bool ignoreCase = false,
        [Description("Host-admitted older-PDF password environment variable.")] string? passwordEnvironmentVariable = null,
        [Description("Host-admitted newer-PDF password environment variable. Null reuses the older password.")] string? comparisonPasswordEnvironmentVariable = null,
        [Description("Explicit permission to replace existing outputs.")] bool overwrite = false,
        [Description("Maximum input bytes per document, 1-268435456.")] long maximumInputBytes = 268435456,
        [Description("Maximum total output bytes, 1-536870912.")] long maximumOutputBytes = 100663296,
        [Description("Maximum serialized result characters, 512-64000.")] int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) => ExecuteSafePdfAsync(() => _service.RedlineRegisteredPdfAsync(sourceId, comparisonSourceId,
            outputPath, changesPath, reportPath, Settings(null, passwordEnvironmentVariable, overwrite, 100, 150, maximumInputBytes, maximumOutputBytes),
            comparisonPasswordEnvironmentVariable, ignoreCase, algorithm, maxOutputCharacters, cancellationToken),
            result => result.Summary ?? "PDF redline: " + result.Status + "; " + result.ArtifactCount + " artifact(s).", result => result.Succeeded);
}
