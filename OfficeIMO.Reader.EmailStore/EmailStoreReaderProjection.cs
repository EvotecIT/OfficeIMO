using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.EmailStore;

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

    internal static OfficeDocumentReadResult EnrichResult(
        OfficeDocumentReadResult result,
        EmailStoreProjection projection) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (projection == null) throw new ArgumentNullException(nameof(projection));
        string storeFormat = projection.ReadResult.Store.Format.ToString();
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
            CreateMetadata("email-store-folder-count", "FolderCount", projection.ReadResult.Store.Folders.Count, "count"),
            CreateMetadata("email-store-item-count", "ItemCount", projection.ReadResult.Store.ItemCount, "count"),
            CreateMetadata("email-store-associated-item-count", "AssociatedItemCount", projection.AssociatedItemCount, "count"),
            CreateMetadata("email-store-diagnostic-count", "DiagnosticCount", projection.Diagnostics.Count, "count"),
            CreateMetadata("email-store-bytes-read", "BytesRead", projection.ReadResult.BytesRead, "number")
        }).ToArray();
        if (!string.IsNullOrWhiteSpace(projection.ReadResult.Store.DisplayName)) {
            result.Source.Title = projection.ReadResult.Store.DisplayName;
        }
        return result;
    }

    private static IReadOnlyDictionary<string, string> BuildFolderPaths(
        IReadOnlyList<global::OfficeIMO.Email.Store.EmailStoreFolder> folders,
        List<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        var byId = new Dictionary<string, global::OfficeIMO.Email.Store.EmailStoreFolder>(StringComparer.Ordinal);
        foreach (global::OfficeIMO.Email.Store.EmailStoreFolder folder in folders) {
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
        foreach (global::OfficeIMO.Email.Store.EmailStoreFolder folder in folders) {
            cancellationToken.ThrowIfCancellationRequested();
            if (resolved.ContainsKey(folder.Id)) continue;
            ResolveFolderPath(folder, byId, resolved, diagnostics);
        }
        return resolved;
    }

    private static void ResolveFolderPath(
        global::OfficeIMO.Email.Store.EmailStoreFolder folder,
        IReadOnlyDictionary<string, global::OfficeIMO.Email.Store.EmailStoreFolder> byId,
        IDictionary<string, string> resolved,
        List<EmailDiagnostic> diagnostics) {
        var chain = new List<global::OfficeIMO.Email.Store.EmailStoreFolder>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        global::OfficeIMO.Email.Store.EmailStoreFolder? current = folder;
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
            if (!byId.TryGetValue(current.ParentId, out global::OfficeIMO.Email.Store.EmailStoreFolder? parent)) {
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
            global::OfficeIMO.Email.Store.EmailStoreFolder segment = chain[index];
            string name = EscapePathSegment(segment.Name);
            basePath = basePath.Length == 0 ? name : basePath + "/" + name;
            if (!resolved.ContainsKey(segment.Id)) resolved.Add(segment.Id, basePath);
        }
    }

    private static string BuildItemPath(string sourceName, string folderPath, string itemKind, int itemIndex) {
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

    private static EmailDiagnostic MapDiagnostic(EmailStoreDiagnostic diagnostic) {
        EmailDiagnosticSeverity severity = diagnostic.Severity == EmailStoreDiagnosticSeverity.Error
            ? EmailDiagnosticSeverity.Error
            : diagnostic.Severity == EmailStoreDiagnosticSeverity.Information
                ? EmailDiagnosticSeverity.Information
                : EmailDiagnosticSeverity.Warning;
        return new EmailDiagnostic(diagnostic.Code, diagnostic.Message, severity, diagnostic.Location);
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

internal sealed class EmailStoreProjection {
    internal EmailStoreProjection(
        EmailStoreReadResult readResult,
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat emailFormat,
        int associatedItemCount) {
        ReadResult = readResult;
        Documents = documents;
        LogicalPaths = logicalPaths;
        Diagnostics = diagnostics;
        EmailFormat = emailFormat;
        AssociatedItemCount = associatedItemCount;
    }

    internal EmailStoreReadResult ReadResult { get; }
    internal IReadOnlyList<EmailDocument> Documents { get; }
    internal IReadOnlyList<string?> LogicalPaths { get; }
    internal IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    internal EmailFileFormat EmailFormat { get; }
    internal int AssociatedItemCount { get; }
}
