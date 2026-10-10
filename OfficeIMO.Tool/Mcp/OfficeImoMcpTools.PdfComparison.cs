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
}
