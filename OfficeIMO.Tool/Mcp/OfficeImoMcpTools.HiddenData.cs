using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Mcp;

internal sealed partial class OfficeImoMcpTools {
    [McpServerTool(
        Name = "officeimo_inspect_hidden_data",
        Title = "Find metadata and hidden data before sharing",
        ReadOnly = true,
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        UseStructuredContent = true,
        OutputSchemaType = typeof(HiddenDataResultDto))]
    [Description("Report what would leave with an Office, Visio, OpenDocument, EPUB, PDF or image file: identity properties, comments, tracked changes, hidden content, speaker notes, custom XML, attachments, bookmarks, scripts and other active content, image GPS and other metadata, and provenance such as Content Credentials and AI-generated declarations. Nothing is changed. Finding values are short previews of document content.")]
    public async Task<CallToolResult> InspectHiddenDataAsync(
        [Description("Path to one local document.")] string path,
        CancellationToken cancellationToken = default) =>
        await ExecuteAsync(
            () => _service.InspectHiddenDataAsync(path, cancellationToken),
            result => result.Format + ": " + result.Findings.Count + " finding(s).").ConfigureAwait(false);

    [McpServerTool(
        Name = "officeimo_remove_hidden_data",
        Title = "Write a copy without metadata and hidden data",
        ReadOnly = false,
        Destructive = true,
        Idempotent = true,
        OpenWorld = false,
        UseStructuredContent = true,
        OutputSchemaType = typeof(HiddenDataResultDto))]
    [Description("Write a cleaned copy of an Office, Visio, OpenDocument, EPUB, PDF or image file and report every finding with its outcome. The source is never changed. The default kinds remove scripts and other active content but keep tracked changes, hidden content, attachments, bookmarks and provenance; add 'provenance' to remove Content Credentials and AI-generated declarations. Package-level Content Credentials and Content Credentials embedded in a PDF are always removed because a changed document cannot keep them valid. Signed PDFs are refused.")]
    public async Task<CallToolResult> RemoveHiddenDataAsync(
        [Description("Path to one local document.")] string path,
        [Description("Destination file path with the same extension. Existing files are protected unless overwrite is true.")] string outputPath,
        [Description("Categories to remove: 'default', 'all', or names such as 'provenance', 'hidden-content', 'tracked-changes', 'comments', 'image-metadata', 'attachments', 'active-content', 'bookmarks'. Omit for default.")] string[]? kinds = null,
        [Description("What to do with tracked changes: keep, accept or reject.")] string trackedChanges = "keep",
        [Description("Provenance carriers to keep when 'provenance' is selected: 'c2pa', 'external' or 'ai-source'. For example keep 'c2pa' to drop only AI-generated declarations.")] string[]? keepProvenance = null,
        [Description("Name that replaces authors on comments and revisions that stay.")] string? anonymousAuthor = null,
        [Description("Remove package signatures that cleaning invalidates. Otherwise signed documents are refused.")] bool removeInvalidatedSignatures = false,
        [Description("Whether an existing destination may be replaced.")] bool overwrite = false,
        CancellationToken cancellationToken = default) =>
        await ExecuteAsync(
            () => _service.RemoveHiddenDataAsync(path, outputPath, kinds, trackedChanges, keepProvenance, anonymousAuthor, removeInvalidatedSignatures, overwrite, cancellationToken),
            result => "Wrote " + result.Output + " with " + result.Findings.Count + " finding(s).").ConfigureAwait(false);
}
