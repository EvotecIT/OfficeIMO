using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using System.Security.Cryptography;

namespace OfficeIMO.Workflows;

/// <summary>Read-only portable email reports using the canonical email, graph, HTML and PDF owners.</summary>
public static partial class EmailEvidenceWorkflow {
    /// <summary>Creates an evidence bundle from one persisted EML/MSG/OFT/TNEF message, with a SHA-256 over the same open source.</summary>
    public static EmailEvidenceResult Create(string sourcePath, EmailEvidenceOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sourcePath);
        var effective = options ?? new EmailEvidenceOptions(); effective.Validate();
        using var stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > effective.ReaderOptions.MaxInputBytes)
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), stream.Length, effective.ReaderOptions.MaxInputBytes);
        string fingerprint = HashSource(stream, cancellationToken);
        stream.Position = 0;
        using var read = new EmailDocumentReader(effective.ReaderOptions).Read(stream, Path.GetFileName(sourcePath), cancellationToken);
        if (read.Document.Format == EmailFileFormat.Unknown) throw new InvalidDataException("The source is not a supported email message.");
        if (HashSource(stream, cancellationToken) != fingerprint) throw new IOException("The email source changed during evidence projection.");
        return Build(fingerprint, "SourceFileSha256", read.Document.Format.ToString(), true,
            new[] { ("message:0", read.Document) }, Array.Empty<EmailEvidenceThreadLink>(), Array.Empty<EmailEvidenceMissingParent>(),
            read.Diagnostics.Select(MapDiagnostic).ToList(), effective, cancellationToken);
    }

    /// <summary>Creates a chronological dossier for the existing graph component containing a selected store item.</summary>
    /// <remarks>Uses the store's durable fingerprint and reports graph truncation and unresolved parent identities. No payload is extracted to disk.</remarks>
    public static EmailEvidenceResult CreateConversation(EmailStoreSession session, EmailStoreItemId selectedItem,
        EmailEvidenceOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session);
        var effective = options ?? new EmailEvidenceOptions(); effective.Validate();
        string fingerprint = session.GetDurableSourceFingerprint(cancellationToken);
        var graph = session.BuildConversationGraph(new EmailConversationGraphOptions(maxItems: effective.MaxItemsScanned,
            maxEdges: Math.Min(100_000, effective.MaxItemsScanned * 10)), cancellationToken);
        var conversation = graph.GetConversation(selectedItem);
        if (conversation.Nodes.Count > effective.MaxMessages) throw new InvalidDataException("The selected conversation exceeds MaxMessages.");
        var ids = conversation.Nodes.Select(node => node.Reference.Id).ToHashSet(StringComparer.Ordinal);
        var links = conversation.Edges.Select(edge => new EmailEvidenceThreadLink(edge.Source.Reference.Id, edge.Target.Reference.Id,
            edge.Kind.ToString(), edge.Reasons.Select(reason => reason.ToString()).ToArray(), edge.IsHeuristic)).ToArray();
        var missing = graph.OrphanReplies.Where(orphan => ids.Contains(orphan.Child.Reference.Id))
            .Select(orphan => new EmailEvidenceMissingParent(orphan.Child.Reference.Id, Clip(orphan.ParentMessageId), orphan.Reason.ToString())).ToArray();
        var diagnostics = session.Diagnostics.Concat(graph.Diagnostics).Select(value =>
            new EmailEvidenceDiagnostic(Clip(value.Code), value.Severity.ToString(), Clip(value.Message), ClipNullable(value.Location))).ToList();
        int diagnosticCursor = session.Diagnostics.Count;
        var result = Build(fingerprint, "StoreDurableSha256", session.Format.ToString(), graph.IsComplete,
            ReadMessages(), links, missing, diagnostics, effective, cancellationToken,
            () => session.Diagnostics.Skip(diagnosticCursor).Select(value => new EmailEvidenceDiagnostic(
                Clip(value.Code), value.Severity.ToString(), Clip(value.Message), ClipNullable(value.Location))));
        if (session.GetDurableSourceFingerprint(cancellationToken) != fingerprint)
            throw new IOException("The store source changed during conversation projection.");
        return result;

        IEnumerable<(string Id, EmailDocument Document)> ReadMessages() {
            foreach (var node in conversation.Nodes.OrderBy(node => node.Summary.SentAt ?? node.Summary.ReceivedAt)
                .ThenBy(node => node.Reference.Id, StringComparer.Ordinal)) {
                cancellationToken.ThrowIfCancellationRequested();
                var item = session.ReadItem(node.Reference, new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies | EmailStoreItemReadParts.Recipients |
                    EmailStoreItemReadParts.AttachmentMetadata, effective.ReaderOptions.MaxDecodedPropertyBytes), cancellationToken);
                yield return (node.Reference.Id, item.Document);
            }
        }
    }

    private static string HashSource(Stream stream, CancellationToken token) {
        stream.Position = 0;
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[64 * 1024];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) { token.ThrowIfCancellationRequested(); hash.AppendData(buffer, 0, read); }
        token.ThrowIfCancellationRequested();
        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }
    private static EmailEvidenceDiagnostic MapDiagnostic(EmailDiagnostic value) =>
        new(Clip(value.Code), value.Severity.ToString(), Clip(value.Message), ClipNullable(value.Location));
    private static string Clip(string value) => value.Length <= 1024 ? value :
        value.Substring(0, char.IsHighSurrogate(value[1023]) ? 1023 : 1024) + "…";
    private static string? ClipNullable(string? value) => value == null ? null : Clip(value);
}
