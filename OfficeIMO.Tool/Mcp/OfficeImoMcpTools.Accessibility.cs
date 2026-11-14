using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Mcp;

internal sealed partial class OfficeImoMcpTools {
    [McpServerTool(
        Name = "officeimo_check_accessibility",
        Title = "Check a Word, Excel or PowerPoint file for accessibility issues",
        ReadOnly = true,
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        UseStructuredContent = true,
        OutputSchemaType = typeof(AgentAccessibilityResult))]
    [Description("Inspect a Word, Excel or PowerPoint file against the accessibility rules of the Office Accessibility Checker: missing alternative text, table header rows, document title and language, headings, unclear hyperlink text, low text contrast, slide titles, merged cells, default or hidden sheets, and blank cells or paragraphs used for spacing. All three formats share the same finding codes and severities. Nothing is changed. Optionally returns a SARIF 2.1.0 log.")]
    public async Task<CallToolResult> CheckAccessibilityAsync(
        [Description("Path to one local Word, Excel or PowerPoint file.")] string path,
        [Description("Use the strict profile: structural warnings become errors and a document title is required. Heuristic findings stay warnings.")] bool strict = false,
        [Description("Also return the full report as a SARIF 2.1.0 log in the 'sarif' field.")] bool includeSarif = false,
        [Description("Maximum findings to return, from 1 through 5000. Counts always cover every finding.")] int maxFindings = OfficeImoAgentService.DefaultAccessibilityFindings,
        CancellationToken cancellationToken = default) =>
        await ExecuteAsync(
            () => _service.InspectAccessibilityAsync(path, strict, includeSarif, maxFindings, cancellationToken),
            result => result.Format + ": " + result.ErrorCount + " error(s), " + result.WarningCount + " warning(s).").ConfigureAwait(false);
}
