using OfficeIMO.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private static AgentInspectResult CreateDocumentInspectResult(
        AgentSourceRegistration source,
        OfficeDocumentReadResult document) {
        return new AgentInspectResult {
            SourceId = source.SourceId,
            Path = source.Path,
            Kind = document.Kind.ToString(),
            Format = document.CapabilitiesUsed.FirstOrDefault(),
            LengthBytes = source.LengthBytes,
            Title = AgentJson.Limit(document.Source.Title, 256),
            Author = AgentJson.Limit(document.Source.Author, 256),
            Subject = AgentJson.Limit(document.Source.Subject, 256),
            Preview = AgentJson.Limit(CreatePreview(document), 480),
            ChunkCount = document.Chunks.Count,
            BlockCount = document.Blocks.Count,
            PageCount = document.Pages.Count,
            TableCount = document.Tables.Count,
            AssetCount = document.Assets.Count,
            MetadataCount = document.Metadata.Count,
            DiagnosticCount = document.Diagnostics.Count,
            Metadata = document.Metadata
                .Where(entry => !string.IsNullOrWhiteSpace(entry.Name))
                .Take(8)
                .Select(entry => new AgentMetadataSummary {
                    Name = AgentJson.Limit(entry.Name, 96),
                    Value = AgentJson.Limit(entry.Value, 192)
                })
                .ToList(),
            Diagnostics = document.Diagnostics
                .Take(5)
                .Select(MapDiagnostic)
                .ToList()
        };
    }

    private static AgentSearchResult SearchDocument(
        AgentSourceRegistration source,
        OfficeDocumentReadResult document,
        string query,
        int take,
        int cursor) {
        int requested = checked(cursor + take + 1);
        OfficeDocumentSearchResult search = document.Search(
            query,
            new OfficeDocumentSearchOptions { MaximumResults = requested });
        OfficeDocumentSearchHit[] page = search.Hits.Skip(cursor).Take(take + 1).ToArray();
        bool hasMore = page.Length > take || search.MaximumResultsReached;
        if (page.Length > take) page = page.Take(take).ToArray();
        var hits = page.Select(hit => new AgentSearchHit {
            Id = AgentOpaqueId.Encode("block", hit.Block.Id),
            Title = AgentJson.Limit(
                hit.Block.Location.HeadingPath ?? hit.Block.Kind,
                160),
            Snippet = CreateSnippet(hit.Block.Text, hit.StartIndex, hit.Length, 280),
            Pages = hit.Pages
                .Where(location => location.Number.HasValue)
                .Select(location => location.Number!.Value)
                .Distinct()
                .OrderBy(number => number)
                .ToArray()
        }).ToList();
        return new AgentSearchResult {
            SourceId = source.SourceId,
            Query = AgentJson.Limit(query, 256),
            Returned = hits.Count,
            NextCursor = hasMore ? cursor + hits.Count : null,
            Truncated = hasMore,
            Results = hits
        };
    }

    private static AgentFetchResult FetchDocument(
        AgentSourceRegistration source,
        OfficeDocumentReadResult document,
        string kind,
        string value,
        string id,
        int cursor) {
        string content;
        string? title;
        string resultKind;
        switch (kind) {
            case "block":
                OfficeDocumentBlock block = document.Blocks.FirstOrDefault(candidate =>
                    string.Equals(candidate.Id, value, StringComparison.Ordinal))
                    ?? throw new AgentUsageException("The requested block no longer exists in the source.");
                content = block.Text ?? string.Empty;
                title = block.Location.HeadingPath ?? block.Kind;
                resultKind = block.Kind;
                break;
            case "chunk":
                ReaderChunk chunk = document.Chunks.FirstOrDefault(candidate =>
                    string.Equals(candidate.Id, value, StringComparison.Ordinal))
                    ?? throw new AgentUsageException("The requested chunk no longer exists in the source.");
                content = chunk.Markdown ?? chunk.Text ?? string.Empty;
                title = chunk.Location.HeadingPath ?? chunk.Location.SourceBlockKind;
                resultKind = chunk.Kind.ToString();
                break;
            default:
                throw new AgentUsageException("The result id is not fetchable as document content.");
        }
        if (cursor > content.Length) {
            throw new AgentUsageException("Fetch cursor is beyond the available content.");
        }
        return new AgentFetchResult {
            SourceId = source.SourceId,
            Id = id,
            Kind = resultKind,
            Title = AgentJson.Limit(title, 192),
            Content = content.Substring(cursor),
            ContentLength = content.Length,
            Metadata = document.Metadata
                .Take(8)
                .Select(entry => new AgentMetadataSummary {
                    Name = AgentJson.Limit(entry.Name, 96),
                    Value = AgentJson.Limit(entry.Value, 192)
                })
                .ToList(),
            Diagnostics = document.Diagnostics.Take(5).Select(MapDiagnostic).ToList()
        };
    }

    private static string CreatePreview(OfficeDocumentReadResult document) {
        if (!string.IsNullOrWhiteSpace(document.Markdown)) return document.Markdown!;
        return string.Join(
            "\n\n",
            document.Chunks
                .Select(chunk => chunk.Markdown ?? chunk.Text)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Take(3));
    }

    private static string CreateSnippet(string text, int start, int length, int maximumCharacters) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        int context = Math.Max(0, (maximumCharacters - length) / 2);
        int snippetStart = Math.Max(0, start - context);
        int snippetEnd = Math.Min(text.Length, start + length + context);
        string prefix = snippetStart > 0 ? "…" : string.Empty;
        string suffix = snippetEnd < text.Length ? "…" : string.Empty;
        return prefix + text.Substring(snippetStart, snippetEnd - snippetStart).Trim() + suffix;
    }

    private static AgentDiagnosticSummary MapDiagnostic(OfficeDocumentDiagnostic diagnostic) =>
        new() {
            Code = AgentJson.Limit(diagnostic.Code, 96),
            Severity = diagnostic.Severity.ToString(),
            Message = AgentJson.Limit(diagnostic.Message, 256)
        };
}
