using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

internal static class EmailStoreReaderProjection {
    internal static EmailStoreProjection Create(
        EmailStoreReadResult readResult,
        string sourceName,
        CancellationToken cancellationToken) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var documents = new List<EmailDocument>();
        var logicalPaths = new List<string?>();
        var diagnostics = readResult.Diagnostics.Select(MapDiagnostic).ToList();
        IReadOnlyDictionary<string, string> folderPaths = BuildFolderPaths(
            readResult.Store.Folders, diagnostics, cancellationToken);
        int itemIndex = 0;
        int associatedItemCount = 0;

        foreach (global::OfficeIMO.Email.Store.EmailStoreFolder folder in readResult.Store.Folders) {
            cancellationToken.ThrowIfCancellationRequested();
            string folderPath = folderPaths.TryGetValue(folder.Id, out string? value)
                ? value
                : EscapePathSegment(folder.Name);
            foreach (EmailStoreItem item in folder.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                documents.Add(item.Document);
                logicalPaths.Add(BuildItemPath(sourceName, folderPath, "item", itemIndex++));
            }
            foreach (EmailStoreItem item in folder.AssociatedItems) {
                cancellationToken.ThrowIfCancellationRequested();
                documents.Add(item.Document);
                logicalPaths.Add(BuildItemPath(sourceName, folderPath, "associated", itemIndex++));
                associatedItemCount++;
            }
        }

        return new EmailStoreProjection(
            readResult,
            documents,
            logicalPaths,
            diagnostics,
            ResolveEmailFormat(documents),
            associatedItemCount);
    }

    internal static EmailStoreProjection Create(
        EmailStoreSession session,
        string sourceName,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        if (session == null) throw new ArgumentNullException(nameof(session));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        if (adapterOptions == null) throw new ArgumentNullException(nameof(adapterOptions));

        EmailStoreInspectionReport inspection = session.Inspect();
        var documents = new List<EmailDocument>();
        var logicalPaths = new List<string?>();
        var itemDiagnostics = new List<EmailDiagnostic>();
        IReadOnlyDictionary<string, string> folderPaths = BuildFolderPaths(
            session.Folders.Select(folder => new FolderPathNode(
                folder.Id, folder.ParentId, folder.Name)).ToArray(),
            itemDiagnostics,
            cancellationToken);
        int probeLimit = adapterOptions.MaxItems == int.MaxValue
            ? int.MaxValue
            : adapterOptions.MaxItems + 1;
        EmailStoreQuery? query = adapterOptions.Query == null
            ? null
            : CopyQuery(adapterOptions.Query, Math.Min(adapterOptions.Query.MaxResults, probeLimit));
        IEnumerable<EmailStoreItemReference> references = query == null
            ? session.EnumerateItems(new EmailStoreEnumerationOptions(
                includeAssociatedItems: GetStoreOptions(adapterOptions).IncludeAssociatedItems,
                includeOrphanedItems: GetStoreOptions(adapterOptions).IncludeOrphanedItems,
                maxItems: probeLimit), cancellationToken)
            : session.Search(query, cancellationToken).Select(result => result.Reference);
        int attempted = 0;
        int associatedItemCount = 0;
        bool selectionLimitReached = false;
        foreach (EmailStoreItemReference reference in references) {
            cancellationToken.ThrowIfCancellationRequested();
            if (attempted >= adapterOptions.MaxItems) {
                selectionLimitReached = true;
                break;
            }
            attempted++;
            try {
                EmailStoreItem item = session.ReadItem(reference,
                    GetItemReadOptions(adapterOptions), cancellationToken);
                string folderPath = folderPaths.TryGetValue(reference.FolderId, out string? value)
                    ? value
                    : "_unknown-folder";
                string kind = reference.IsAssociated ? "associated" : reference.IsOrphaned ? "recovered" : "item";
                documents.Add(item.Document);
                logicalPaths.Add(BuildItemPath(sourceName, folderPath, kind, documents.Count - 1));
                if (reference.IsAssociated) associatedItemCount++;
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is NotSupportedException ||
                exception is KeyNotFoundException ||
                exception is EmailStoreLimitExceededException) {
                if (!adapterOptions.ContinueOnItemError) throw;
                itemDiagnostics.Add(new EmailDiagnostic(
                    "EMAIL_STORE_READER_ITEM_SKIPPED",
                    exception.Message,
                    exception is EmailStoreLimitExceededException
                        ? EmailDiagnosticSeverity.Warning
                        : EmailDiagnosticSeverity.Error,
                    string.Concat("item/", reference.Id)));
            }
        }
        if (query != null && attempted >= query.MaxResults) selectionLimitReached = true;
        if (selectionLimitReached) {
            itemDiagnostics.Add(new EmailDiagnostic(
                "EMAIL_STORE_READER_SELECTION_LIMIT",
                "Reader stopped at the configured email-store item or query result bound.",
                EmailDiagnosticSeverity.Warning,
                sourceName));
        }
        IReadOnlyList<EmailDiagnostic> diagnostics = session.Diagnostics
            .Select(MapDiagnostic)
            .Concat(itemDiagnostics)
            .ToArray();
        return new EmailStoreProjection(
            session.Format,
            session.DisplayName,
            session.SourceLength,
            inspection.FolderCount,
            inspection.DeclaredItemCount,
            inspection.FoldersWithUnknownItemCount,
            documents,
            logicalPaths,
            diagnostics,
            ResolveEmailFormat(documents),
            associatedItemCount,
            selectionLimitReached);
    }

    internal static OfficeDocumentReadResult EnrichResult(
        OfficeDocumentReadResult result,
        EmailStoreProjection projection) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (projection == null) throw new ArgumentNullException(nameof(projection));
        string storeFormat = projection.StoreFormat.ToString();
        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(new[] {
                OfficeDocumentReaderBuilderEmailStoreExtensions.HandlerId,
                "officeimo.email.store",
                "officeimo.email.store." + storeFormat.ToLowerInvariant()
            })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        result.Metadata = result.Metadata.Concat(new[] {
            CreateMetadata("email-store-format", "StoreFormat", storeFormat, "string"),
            CreateMetadata("email-store-folder-count", "FolderCount", projection.FolderCount, "count"),
            CreateMetadata("email-store-item-count", "ItemCount", projection.Documents.Count, "count"),
            CreateMetadata("email-store-associated-item-count", "AssociatedItemCount", projection.AssociatedItemCount, "count"),
            CreateMetadata("email-store-diagnostic-count", "DiagnosticCount", projection.Diagnostics.Count, "count"),
            CreateMetadata("email-store-source-length", "SourceLength", projection.SourceLength, "number"),
            CreateMetadata("email-store-declared-item-count", "DeclaredItemCount", projection.DeclaredItemCount, "count"),
            CreateMetadata("email-store-unknown-item-count-folders", "FoldersWithUnknownItemCount",
                projection.FoldersWithUnknownItemCount, "count"),
            CreateMetadata("email-store-selection-limit-reached", "SelectionLimitReached",
                projection.SelectionLimitReached, "boolean")
        }).ToArray();
        if (!string.IsNullOrWhiteSpace(projection.DisplayName)) {
            result.Source.Title = projection.DisplayName;
        }
        return result;
    }

    private static IReadOnlyDictionary<string, string> BuildFolderPaths(
        IReadOnlyList<global::OfficeIMO.Email.Store.EmailStoreFolder> folders,
        List<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken) =>
        BuildFolderPaths(folders.Select(folder => new FolderPathNode(
            folder.Id, folder.ParentId, folder.Name)).ToArray(), diagnostics, cancellationToken);

    internal static IReadOnlyDictionary<string, string> BuildFolderPaths(
        IReadOnlyList<FolderPathNode> folders,
        List<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        var byId = new Dictionary<string, FolderPathNode>(StringComparer.Ordinal);
        foreach (FolderPathNode folder in folders) {
            if (!byId.ContainsKey(folder.Id)) {
                byId.Add(folder.Id, folder);
            } else {
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_STORE_DUPLICATE_FOLDER_ID",
                    "The email store contains more than one folder with the same identifier.",
                    EmailDiagnosticSeverity.Warning,
                    folder.Id));
            }
        }

        var resolved = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (FolderPathNode folder in folders) {
            cancellationToken.ThrowIfCancellationRequested();
            if (resolved.ContainsKey(folder.Id)) continue;
            ResolveFolderPath(folder, byId, resolved, diagnostics);
        }
        return resolved;
    }

    private static void ResolveFolderPath(
        FolderPathNode folder,
        IReadOnlyDictionary<string, FolderPathNode> byId,
        IDictionary<string, string> resolved,
        List<EmailDiagnostic> diagnostics) {
        var chain = new List<FolderPathNode>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        FolderPathNode? current = folder;
        string basePath = string.Empty;
        while (current != null) {
            if (resolved.TryGetValue(current.Id, out string? cached)) {
                basePath = cached;
                break;
            }
            if (!visited.Add(current.Id)) {
                basePath = "_invalid-hierarchy";
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_STORE_FOLDER_CYCLE",
                    "A cycle in the email-store folder hierarchy was isolated.",
                    EmailDiagnosticSeverity.Warning,
                    current.Id));
                break;
            }
            chain.Add(current);
            if (current.ParentId == null) break;
            if (!byId.TryGetValue(current.ParentId, out FolderPathNode? parent)) {
                basePath = "_missing-parent";
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_STORE_FOLDER_PARENT_MISSING",
                    "A folder references a parent that is not present in the materialized store.",
                    EmailDiagnosticSeverity.Warning,
                    current.Id));
                break;
            }
            current = parent;
        }

        for (int index = chain.Count - 1; index >= 0; index--) {
            FolderPathNode segment = chain[index];
            string name = EscapePathSegment(segment.Name);
            basePath = basePath.Length == 0 ? name : basePath + "/" + name;
            if (!resolved.ContainsKey(segment.Id)) resolved.Add(segment.Id, basePath);
        }
    }

    internal static string BuildItemPath(string sourceName, string folderPath, string itemKind, int itemIndex) {
        string prefix = sourceName + "!/";
        if (!string.IsNullOrEmpty(folderPath)) prefix += folderPath + "/";
        return prefix + itemKind + "-" + itemIndex.ToString("D6", CultureInfo.InvariantCulture);
    }

    private static string EscapePathSegment(string? value) {
        string segment = string.IsNullOrWhiteSpace(value) ? "_unnamed" : value!.Trim();
        return segment
            .Replace("%", "%25")
            .Replace("!", "%21")
            .Replace("/", "%2F")
            .Replace("\\", "%5C");
    }

    private static EmailFileFormat ResolveEmailFormat(IReadOnlyList<EmailDocument> documents) {
        EmailFileFormat resolved = EmailFileFormat.Unknown;
        for (int index = 0; index < documents.Count; index++) {
            EmailFileFormat candidate = documents[index].Format;
            if (candidate == EmailFileFormat.Unknown) continue;
            if (resolved == EmailFileFormat.Unknown) {
                resolved = candidate;
            } else if (resolved != candidate) {
                return EmailFileFormat.Unknown;
            }
        }
        return resolved;
    }

    internal static EmailDiagnostic MapDiagnostic(EmailStoreDiagnostic diagnostic) {
        EmailDiagnosticSeverity severity = diagnostic.Severity == EmailStoreDiagnosticSeverity.Error
            ? EmailDiagnosticSeverity.Error
            : diagnostic.Severity == EmailStoreDiagnosticSeverity.Information
                ? EmailDiagnosticSeverity.Information
                : EmailDiagnosticSeverity.Warning;
        return new EmailDiagnostic(diagnostic.Code, diagnostic.Message, severity, diagnostic.Location);
    }

    internal static EmailStoreQuery CopyQuery(EmailStoreQuery query, int maxResults) =>
        new EmailStoreQuery(
            query.FolderId,
            query.IncludeDescendants,
            query.IncludeAssociatedItems,
            query.IncludeOrphanedItems,
            query.ItemKind,
            query.SubjectContains,
            query.SenderContains,
            query.Since,
            query.Before,
            query.HasAttachments,
            query.IsRead,
            query.MaxItemsScanned,
            maxResults);

    internal static EmailStoreReaderOptions GetStoreOptions(ReaderEmailStoreOptions options) =>
        options.StoreOptions ?? EmailStoreReaderOptions.Default;

    internal static EmailStoreItemReadOptions GetItemReadOptions(ReaderEmailStoreOptions options) {
        EmailStoreItemReadOptions selected = options.ItemReadOptions ?? EmailStoreItemReadOptions.Default;
        return options.StreamAttachmentContent && !selected.PreferStreamingAttachmentContent
            ? new EmailStoreItemReadOptions(
                selected.Parts, selected.MaxDecodedPropertyBytes,
                preferStreamingAttachmentContent: true)
            : selected;
    }

    private static OfficeDocumentMetadataEntry CreateMetadata(
        string id,
        string name,
        object value,
        string valueType) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "email.store.summary",
            Name = name,
            Value = Convert.ToString(value, CultureInfo.InvariantCulture),
            ValueType = valueType
        };
    }
}

internal sealed class FolderPathNode {
    internal FolderPathNode(string id, string? parentId, string name) {
        Id = id;
        ParentId = parentId;
        Name = name;
    }

    internal string Id { get; }
    internal string? ParentId { get; }
    internal string Name { get; }
}

internal sealed class EmailStoreProjection {
    internal EmailStoreProjection(
        EmailStoreReadResult readResult,
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat emailFormat,
        int associatedItemCount) {
        StoreFormat = readResult.Store.Format;
        DisplayName = readResult.Store.DisplayName;
        SourceLength = readResult.BytesRead;
        FolderCount = readResult.Store.Folders.Count;
        DeclaredItemCount = readResult.Store.ItemCount;
        FoldersWithUnknownItemCount = 0;
        Documents = documents;
        LogicalPaths = logicalPaths;
        Diagnostics = diagnostics;
        EmailFormat = emailFormat;
        AssociatedItemCount = associatedItemCount;
    }

    internal EmailStoreProjection(
        EmailStoreFormat storeFormat,
        string? displayName,
        long sourceLength,
        int folderCount,
        long declaredItemCount,
        int foldersWithUnknownItemCount,
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat emailFormat,
        int associatedItemCount,
        bool selectionLimitReached) {
        StoreFormat = storeFormat;
        DisplayName = displayName;
        SourceLength = sourceLength;
        FolderCount = folderCount;
        DeclaredItemCount = declaredItemCount;
        FoldersWithUnknownItemCount = foldersWithUnknownItemCount;
        Documents = documents;
        LogicalPaths = logicalPaths;
        Diagnostics = diagnostics;
        EmailFormat = emailFormat;
        AssociatedItemCount = associatedItemCount;
        SelectionLimitReached = selectionLimitReached;
    }

    internal EmailStoreFormat StoreFormat { get; }
    internal string? DisplayName { get; }
    internal long SourceLength { get; }
    internal int FolderCount { get; }
    internal long DeclaredItemCount { get; }
    internal int FoldersWithUnknownItemCount { get; }
    internal IReadOnlyList<EmailDocument> Documents { get; }
    internal IReadOnlyList<string?> LogicalPaths { get; }
    internal IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    internal EmailFileFormat EmailFormat { get; }
    internal int AssociatedItemCount { get; }
    internal bool SelectionLimitReached { get; }
}
