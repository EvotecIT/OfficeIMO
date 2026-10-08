using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal Task<AgentEmailSearchResult> SearchEmailContentAsync(string path, string query,
        string fields = "All", string? checkpoint = null, int take = 10, int maxItemsScanned = 10_000,
        long maxDecodedBytes = 16L * 1024 * 1024, int maxSearchableCharacters = 2_000_000,
        string? subject = null, string? sender = null, string? folderId = null,
        DateTimeOffset? since = null, DateTimeOffset? before = null, bool? hasAttachments = null,
        bool? isRead = null, bool includeDescendants = false,
        int maxOutputCharacters = DefaultSearchOutputCharacters, CancellationToken cancellationToken = default) {
        ValidateSearchBounds(take, 0);
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        if (string.IsNullOrWhiteSpace(query) || query.Length > 1024)
            throw new AgentUsageException("Email content search requires a query of 1 through 1024 characters.");
        if (maxItemsScanned < 1 || maxItemsScanned > 10_000 || maxDecodedBytes < 1 || maxDecodedBytes > 64L * 1024 * 1024 ||
            maxSearchableCharacters < 1 || maxSearchableCharacters > 2_000_000)
            throw new AgentUsageException("Email search bounds are 1-10000 scanned items, 1-67108864 decoded bytes and 1-2000000 searchable characters per item.");
        if (!Enum.TryParse(fields, ignoreCase: true, out EmailStoreContentSearchFields selectedFields) ||
            selectedFields == EmailStoreContentSearchFields.None || (selectedFields & ~EmailStoreContentSearchFields.All) != 0)
            throw new AgentUsageException("Email search fields must name one or more supported semantic fields.");
        EmailStoreContentSearchCheckpoint? resume = null;
        if (checkpoint != null && !EmailStoreContentSearchCheckpoint.TryParse(checkpoint, out resume))
            throw new AgentUsageException("The email search checkpoint is invalid.");
        string inputPath = _pathPolicy.ResolveInput(path);
        var storeOptions = CreateEmailStoreOptions(maxDecodedBytes);
        var source = _registry.RegisterEmailStore(inputPath, storeOptions, cancellationToken);
        if (!IsEmailStoreSource(source)) throw new AgentUsageException("Email content search requires a mailbox file or directory.");
        using var session = EmailStoreSession.Open(source.Path, storeOptions, cancellationToken);
        EmailStoreContentSearchReport report;
        var safety = new AgentContentSafetySummary();
        try {
            report = session.SearchContent(new EmailStoreContentQuery(new[] { query }, fields: selectedFields,
                metadataFilter: new EmailStoreQuery(folderId: folderId, includeDescendants: includeDescendants,
                    subjectContains: subject, senderContains: sender, since: since, before: before,
                    hasAttachments: hasAttachments, isRead: isRead),
                maxItemsScanned: maxItemsScanned, maxResults: take,
                maxDecodedPropertyBytesPerItem: maxDecodedBytes, maxSearchableCharactersPerItem: maxSearchableCharacters,
                resumeFrom: resume, bodyTextProjector: new EmailStoreHtmlBodyTextProjector(
                    EmailConcealedTextPolicy.ExcludeRemovable, safety.Include)), cancellationToken: cancellationToken);
        } catch (ArgumentException exception) {
            throw new AgentUsageException(exception.Message);
        }
        var hits = report.Results.Select(match => new AgentSearchHit {
            Id = AgentOpaqueId.Encode("mail", match.Reference.Id),
            Title = AgentJson.Limit(match.Summary.Subject, 192), Snippet = AgentJson.Limit(match.Snippet, 320),
            Sender = AgentJson.Limit(match.Summary.From?.ToString() ?? match.Summary.Sender?.ToString(), 192),
            Timestamp = match.Summary.ReceivedAt ?? match.Summary.SentAt, FolderId = match.Reference.FolderId,
            MatchedFields = match.MatchedFields.ToString()
        }).ToList();
        var diagnostics = session.Diagnostics.Concat(report.Diagnostics).Take(5).Select(value => new AgentDiagnosticSummary {
            Code = AgentJson.Limit(value.Code, 96), Severity = value.Severity.ToString(), Message = AgentJson.Limit(value.Message, 256)
        }).ToList();
        var result = new AgentEmailSearchResult {
            ContentSafety = safety,
            SourceId = source.SourceId, Returned = hits.Count, ItemsScanned = report.ItemsScanned, ItemsSkipped = report.ItemsSkipped,
            ScanLimitReached = report.StoppedAtItemLimit, IsComplete = report.IsComplete, Truncated = !report.IsComplete,
            NextCheckpoint = report.NextCheckpoint?.Value, DiagnosticCount = session.Diagnostics.Count + report.Diagnostics.Count,
            Diagnostics = diagnostics, Results = hits
        };
        while (AgentJson.Measure(result) > maxOutputCharacters) {
            result.Truncated = true;
            var verbose = hits.FirstOrDefault(hit => !string.IsNullOrEmpty(hit.Title) || !string.IsNullOrEmpty(hit.Snippet) ||
                !string.IsNullOrEmpty(hit.Sender) || !string.IsNullOrEmpty(hit.FolderId) || !string.IsNullOrEmpty(hit.MatchedFields));
            if (verbose != null) {
                verbose.Title = null; verbose.Snippet = null; verbose.Sender = null; verbose.FolderId = null; verbose.MatchedFields = null;
            } else if (diagnostics.Count > 0) diagnostics.RemoveAt(diagnostics.Count - 1);
            else if (hits.Count > 1) {
                hits.RemoveAt(hits.Count - 1); result.Returned = hits.Count; result.IsComplete = false;
                result.NextCheckpoint = report.Results[hits.Count - 1].ResumeAfter.Value;
            } else if (hits.Count == 1) {
                // Keep the input position: this match was scanned but never delivered.
                hits.Clear(); result.Returned = 0; result.IsComplete = false; result.NextCheckpoint = checkpoint;
                result.DiagnosticCount++;
                diagnostics.Add(new AgentDiagnosticSummary { Code = "EMAIL_SEARCH_OUTPUT_BUDGET", Severity = "Warning",
                    Message = "A match cannot fit the output budget. Retry this position with a larger budget or use the store API." });
            } else break;
        }
        EnsureWithinBudget(result, maxOutputCharacters);
        return Task.FromResult(result);
    }
}
