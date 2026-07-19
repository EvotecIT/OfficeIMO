using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

/// <summary>Item-at-a-time Reader ingestion for bounded PST, OST, OLM, EMLX, and mailbox-directory sessions.</summary>
public static class EmailStoreItemReader {
    /// <summary>
    /// Lazily selects and projects one store item at a time. The source session and deferred attachment streams
    /// remain open only for the lifetime of the returned enumeration.
    /// </summary>
    public static IEnumerable<ReaderEmailStoreItemResult> Read(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailStoreOptions? emailStoreOptions = null,
        CancellationToken cancellationToken = default) {
        return Read(OfficeDocumentReader.Default, path, readerOptions,
            emailStoreOptions, cancellationToken);
    }

    /// <summary>
    /// Lazily selects and projects one store item at a time through the supplied Reader's configured
    /// HTML, RTF, attachment, and other modular handlers.
    /// </summary>
    public static IEnumerable<ReaderEmailStoreItemResult> Read(
        OfficeDocumentReader reader,
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailStoreOptions? emailStoreOptions = null,
        CancellationToken cancellationToken = default) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (path == null) throw new ArgumentNullException(nameof(path));
        return reader.Scope(ReadCore(path, readerOptions ?? new ReaderOptions(),
            ReaderEmailStoreOptionsCloner.CloneOrDefault(emailStoreOptions), cancellationToken));
    }

    private static IEnumerable<ReaderEmailStoreItemResult> ReadCore(
        string path,
        ReaderOptions readerOptions,
        ReaderEmailStoreOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailStoreReaderOptions effective = ReaderEmailStoreOptionsCloner.CreateEffective(
            adapterOptions, readerOptions);
        using (EmailStoreSession session = EmailStoreSession.Open(path, effective, cancellationToken)) {
            var hierarchyDiagnostics = new List<EmailDiagnostic>();
            IReadOnlyDictionary<string, string> folderPaths =
                EmailStoreReaderProjection.BuildFolderPaths(
                    session.Folders.Select(folder => new FolderPathNode(
                        folder.Id, folder.ParentId, folder.Name)).ToArray(),
                    hierarchyDiagnostics,
                    cancellationToken);
            EmailDiagnostic[] openingDiagnostics = hierarchyDiagnostics
                .Concat(session.Diagnostics.Select(EmailStoreReaderProjection.MapDiagnostic))
                .ToArray();
            int probeLimit = adapterOptions.MaxItems == int.MaxValue
                ? int.MaxValue
                : adapterOptions.MaxItems + 1;
            EmailStoreQuery? query = adapterOptions.Query == null
                ? null
                : EmailStoreReaderProjection.CopyQuery(adapterOptions.Query,
                    Math.Min(adapterOptions.Query.MaxResults, probeLimit));
            IEnumerable<EmailStoreItemReference> references = query == null
                ? session.EnumerateItems(new EmailStoreEnumerationOptions(
                    includeAssociatedItems: EmailStoreReaderProjection.GetStoreOptions(adapterOptions).IncludeAssociatedItems,
                    includeOrphanedItems: EmailStoreReaderProjection.GetStoreOptions(adapterOptions).IncludeOrphanedItems,
                    maxItems: probeLimit), cancellationToken)
                : session.Search(query, cancellationToken).Select(result => result.Reference);

            var cursor = new EmailDocumentProjectionCursor();
            int diagnosticCursor = session.Diagnostics.Count;
            int attempted = 0;
            foreach (EmailStoreItemReference reference in references) {
                cancellationToken.ThrowIfCancellationRequested();
                if (attempted >= adapterOptions.MaxItems) yield break;
                int itemIndex = attempted++;
                string folderPath = folderPaths.TryGetValue(reference.FolderId, out string? value)
                    ? value
                    : "_unknown-folder";
                string kind = reference.IsAssociated
                    ? "associated"
                    : reference.IsOrphaned ? "recovered" : "item";
                string logicalPath = EmailStoreReaderProjection.BuildItemPath(
                    path, folderPath, kind, itemIndex);
                EmailStoreItemSummary? summary = reference.Summary;
                ReaderEmailStoreItemResult result;
                try {
                    summary = summary ?? session.ReadSummary(reference, cancellationToken);
                    EmailStoreItem item = session.ReadItem(reference,
                        EmailStoreReaderProjection.GetItemReadOptions(adapterOptions), cancellationToken);
                    var itemDiagnostics = new List<EmailDiagnostic>();
                    AddNewStoreDiagnostics(session, itemDiagnostics, ref diagnosticCursor);
                    IReadOnlyList<ReaderChunk> chunks =
                        EmailReaderProjection.ProjectEmailDocumentToChunks(
                            item.Document,
                            logicalPath,
                            itemDiagnostics,
                            path,
                            readerOptions,
                            cursor,
                            cancellationToken);
                    result = new ReaderEmailStoreItemResult(
                        reference, summary, logicalPath, chunks, itemDiagnostics,
                        itemIndex == 0 ? openingDiagnostics : null);
                } catch (Exception exception) when (
                    adapterOptions.ContinueOnItemError &&
                    (exception is InvalidDataException ||
                     exception is NotSupportedException ||
                     exception is KeyNotFoundException ||
                     exception is EmailStoreLimitExceededException)) {
                    var itemDiagnostics = new List<EmailDiagnostic>();
                    AddNewStoreDiagnostics(session, itemDiagnostics, ref diagnosticCursor);
                    itemDiagnostics.Add(new EmailDiagnostic(
                        "EMAIL_STORE_READER_ITEM_SKIPPED",
                        exception.Message,
                        exception is EmailStoreLimitExceededException
                            ? EmailDiagnosticSeverity.Warning
                            : EmailDiagnosticSeverity.Error,
                        string.Concat("item/", reference.Id)));
                    result = new ReaderEmailStoreItemResult(
                        reference, summary, logicalPath,
                        Array.Empty<ReaderChunk>(), itemDiagnostics,
                        itemIndex == 0 ? openingDiagnostics : null);
                }
                yield return result;
            }
        }
    }

    private static void AddNewStoreDiagnostics(EmailStoreSession session,
        ICollection<EmailDiagnostic> target, ref int diagnosticCursor) {
        IReadOnlyList<EmailStoreDiagnostic> diagnostics = session.Diagnostics;
        while (diagnosticCursor < diagnostics.Count) {
            target.Add(EmailStoreReaderProjection.MapDiagnostic(diagnostics[diagnosticCursor++]));
        }
    }
}
