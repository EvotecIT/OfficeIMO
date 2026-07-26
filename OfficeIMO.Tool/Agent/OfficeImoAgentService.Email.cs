using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private static AgentInspectResult InspectEmailStore(
        AgentSourceRegistration source,
        int maximumCharacters,
        CancellationToken cancellationToken) {
        using EmailStoreSession session = EmailStoreSession.Open(
            source.Path,
            CreateEmailStoreOptions(),
            cancellationToken);
        int folderSampleLimit = Math.Min(
            session.Folders.Count,
            Math.Max(1, maximumCharacters / 128));
        var folders = session.Folders
            .Take(folderSampleLimit)
            .Select(folder => new AgentFolderSummary {
                Id = folder.Id,
                ParentId = folder.ParentId,
                Name = AgentJson.Limit(folder.Name, 192),
                ItemCount = folder.ItemCount,
                AssociatedItemCount = folder.AssociatedItemCount,
                SpecialKind = folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Unknown
                    ? null
                    : folder.SpecialFolderKind.ToString()
            })
            .ToList();
        int declaredItems = session.Folders.Sum(folder => folder.ItemCount ?? 0);
        return new AgentInspectResult {
            SourceId = source.SourceId,
            Path = source.Path,
            Kind = "emailStore",
            Format = session.Format.ToString(),
            LengthBytes = source.LengthBytes ?? session.SourceLength,
            Title = AgentJson.Limit(session.DisplayName, 256),
            Preview = session.Folders.Count + " folders; " + declaredItems + " declared items",
            FolderCount = session.Folders.Count,
            DiagnosticCount = session.Diagnostics.Count,
            Folders = folders,
            Truncated = folderSampleLimit < session.Folders.Count,
            Diagnostics = session.Diagnostics
                .Take(5)
                .Select(diagnostic => new AgentDiagnosticSummary {
                    Code = AgentJson.Limit(diagnostic.Code, 96),
                    Severity = diagnostic.Severity.ToString(),
                    Message = AgentJson.Limit(diagnostic.Message, 256)
                })
                .ToList()
        };
    }

    private static AgentSearchResult SearchEmailStore(
        AgentSourceRegistration source,
        string? query,
        string? subject,
        string? sender,
        string? folderId,
        DateTimeOffset? since,
        DateTimeOffset? before,
        bool? hasAttachments,
        bool? isRead,
        bool includeDescendants,
        int take,
        int cursor,
        CancellationToken cancellationToken) {
        string? subjectFilter = string.IsNullOrWhiteSpace(subject) ? query : subject;
        using EmailStoreSession session = EmailStoreSession.Open(
            source.Path,
            CreateEmailStoreOptions(),
            cancellationToken);
        int requested = checked(cursor + take + 1);
        var emailQuery = new EmailStoreQuery(
            folderId: folderId,
            includeDescendants: includeDescendants,
            subjectContains: subjectFilter,
            senderContains: sender,
            since: since,
            before: before,
            hasAttachments: hasAttachments,
            isRead: isRead,
            maxItemsScanned: int.MaxValue,
            maxResults: requested);
        EmailStoreSearchResult[] page = session.Search(emailQuery, cancellationToken)
            .Skip(cursor)
            .Take(take + 1)
            .ToArray();
        bool hasMore = page.Length > take;
        if (page.Length > take) page = page.Take(take).ToArray();
        var hits = page.Select(result => new AgentSearchHit {
            Id = AgentOpaqueId.Encode("mail", result.Reference.Id),
            Title = AgentJson.Limit(result.Summary.Subject, 192),
            Snippet = AgentJson.Limit(CreateEmailSummary(result.Summary), 320),
            Sender = AgentJson.Limit(
                result.Summary.From?.ToString() ?? result.Summary.Sender?.ToString(),
                192),
            Timestamp = result.Summary.ReceivedAt ?? result.Summary.SentAt,
            FolderId = result.Reference.FolderId
        }).ToList();
        return new AgentSearchResult {
            SourceId = source.SourceId,
            Query = AgentJson.Limit(query ?? subject, 256),
            Returned = hits.Count,
            NextCursor = hasMore ? cursor + hits.Count : null,
            Truncated = hasMore,
            Results = hits
        };
    }

    private static AgentFetchResult FetchEmailStoreItem(
        AgentSourceRegistration source,
        string itemId,
        string id,
        int cursor,
        CancellationToken cancellationToken) {
        OfficeDocumentReader reader = CreateReader();
        ReaderEmailStoreItemResult item = reader.ReadEmailStoreItem(
            source.Path,
            itemId,
            new ReaderOptions {
                MaxChars = DefaultFetchOutputCharacters,
                ComputeHashes = false
            },
            new ReaderEmailStoreOptions {
                StoreOptions = CreateEmailStoreOptions(),
                ItemReadOptions = new EmailStoreItemReadOptions(
                    AgentEmailParts,
                    preferStreamingAttachmentContent: true),
                MaxItems = 1,
                StreamAttachmentContent = true,
                ComputeSourceHash = false
            },
            cancellationToken);
        string content = string.Join(
            "\n\n",
            item.Chunks
                .Select(chunk => chunk.Markdown ?? chunk.Text)
                .Where(value => !string.IsNullOrWhiteSpace(value)));
        if (cursor > content.Length) {
            throw new AgentUsageException("Fetch cursor is beyond the available content.");
        }
        var metadata = new List<AgentMetadataSummary>();
        AddMetadata(metadata, "subject", item.Summary?.Subject);
        AddMetadata(metadata, "from", item.Summary?.From?.ToString());
        AddMetadata(metadata, "sender", item.Summary?.Sender?.ToString());
        AddMetadata(metadata, "messageId", item.Summary?.MessageId);
        AddMetadata(metadata, "receivedAt", item.Summary?.ReceivedAt?.ToString("O"));
        AddMetadata(metadata, "sentAt", item.Summary?.SentAt?.ToString("O"));
        AddMetadata(metadata, "hasAttachments", item.Summary?.HasAttachments?.ToString());
        AddMetadata(metadata, "folderId", item.Reference.FolderId);
        return new AgentFetchResult {
            SourceId = source.SourceId,
            Id = id,
            Kind = item.Summary?.OutlookItemKind.ToString() ?? "email",
            Title = AgentJson.Limit(item.Summary?.Subject, 192),
            Content = content.Substring(cursor),
            ContentLength = content.Length,
            Metadata = metadata,
            Diagnostics = item.Diagnostics.Take(5).Select(diagnostic => new AgentDiagnosticSummary {
                Code = AgentJson.Limit(diagnostic.Code, 96),
                Severity = diagnostic.Severity.ToString(),
                Message = AgentJson.Limit(diagnostic.Message, 256)
            }).ToList()
        };
    }

    private static EmailStoreReaderOptions CreateEmailStoreOptions() =>
        new(
            retainAttachmentContent: false,
            maxItemCount: 1_000_000,
            maxTotalAttachmentBytes: 64L * 1024 * 1024);

    private static string CreateEmailSummary(EmailStoreItemSummary summary) {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(summary.Subject)) parts.Add(summary.Subject!);
        if (summary.From != null) parts.Add("From: " + summary.From);
        DateTimeOffset? timestamp = summary.ReceivedAt ?? summary.SentAt;
        if (timestamp.HasValue) parts.Add(timestamp.Value.ToString("O"));
        if (summary.HasAttachments == true) parts.Add("Has attachments");
        return string.Join(" | ", parts);
    }

    private static void AddMetadata(
        ICollection<AgentMetadataSummary> metadata,
        string name,
        string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        metadata.Add(new AgentMetadataSummary {
            Name = name,
            Value = AgentJson.Limit(value, 256)
        });
    }
}