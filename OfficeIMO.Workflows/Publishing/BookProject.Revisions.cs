using OfficeIMO.Internal;
using OfficeIMO.Epub;
using System.Security.Cryptography;

namespace OfficeIMO.Workflows;

/// <summary>A named, immutable publication snapshot retained inside a book project.</summary>
public sealed record BookProjectRevision(string Id, string Name, DateTimeOffset CreatedUtc, string Sha256, long PublicationBytes);

public sealed partial class BookProject {
    /// <summary>Maximum combined stored EPUB bytes in named revisions.</summary>
    public const long MaximumRevisionBytes = 128L * 1024 * 1024;
    /// <summary>Maximum combined expanded EPUB bytes in named revisions.</summary>
    public const long MaximumRevisionExpandedBytes = 256L * 1024 * 1024;
    /// <summary>Maximum named revisions retained by one project.</summary>
    public const int MaximumRevisionCount = 100;
    private readonly List<(BookProjectRevision Info, byte[] Bytes, long ExpandedBytes)> _revisions = [];
    private long _revisionExpandedBytes;
    /// <summary>Named revisions in creation order. Import review state is project-wide, not reverted with content.</summary>
    public IReadOnlyList<BookProjectRevision> Revisions => Array.AsReadOnly(_revisions.Select(item => item.Info).ToArray());

    /// <summary>Captures the current validated publication. Exhausted bounds reject capture without evicting earlier revisions.</summary>
    public BookProjectRevision CreateRevision(string name, CancellationToken cancellationToken = default) {
        if ((name == null || string.IsNullOrWhiteSpace(name))) throw new ArgumentException("A non-empty value is required.", nameof(name));
        if (name.Length > 256) throw new ArgumentException("Revision names cannot exceed 256 characters.", nameof(name));
        cancellationToken.ThrowIfCancellationRequested();
        if (_revisions.Count >= MaximumRevisionCount) throw new InvalidOperationException("The project revision-count bound is exhausted.");
        byte[] bytes = _publication.Write(cancellationToken: cancellationToken).Bytes;
        if (bytes.LongLength > MaximumRevisionBytes - _revisions.Sum(item => item.Bytes.LongLength))
            throw new InvalidOperationException("The project revision-byte bound is exhausted.");
        long expanded;
        try {
            expanded = CountPublicationExpandedBytes(bytes, MaximumRevisionExpandedBytes - _revisionExpandedBytes, cancellationToken);
        } catch (InvalidDataException error) {
            throw new InvalidOperationException("The project revision expansion bound is exhausted.", error);
        }
        var info = new BookProjectRevision(Guid.NewGuid().ToString("N"), name, DateTimeOffset.UtcNow, OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(bytes)), bytes.LongLength);
        cancellationToken.ThrowIfCancellationRequested();
        _revisionExpandedBytes += expanded;
        _revisions.Add((info, bytes, expanded));
        return info;
    }

    /// <summary>Restores a named publication snapshot as an undoable edit, retaining all revisions and import review state.</summary>
    public void RestoreRevision(string revisionId, CancellationToken cancellationToken = default) {
        var revision = RequireRevision(revisionId);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] current = _publication.Write(cancellationToken: cancellationToken).Bytes;
        using var stream = new MemoryStream(revision.Bytes, false);
        var restored = EpubPublication.Load(stream, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (!current.SequenceEqual(revision.Bytes)) { _undo = current.LongLength <= MaximumPublicationBytes ? current : null; _redo = null; }
        _publication = restored;
    }

    /// <summary>Compares a retained revision with the current publication through the EPUB owner.</summary>
    public EpubEditionComparison CompareRevision(string revisionId, CancellationToken cancellationToken = default) {
        var revision = RequireRevision(revisionId);
        using var stream = new MemoryStream(revision.Bytes, false);
        var baseline = EpubPublication.Load(stream, cancellationToken: cancellationToken);
        return baseline.CompareTo(_publication, cancellationToken);
    }

    /// <summary>Explicitly removes a named snapshot without changing the current publication or session undo.</summary>
    public void RemoveRevision(string revisionId) {
        var revision = RequireRevision(revisionId);
        _revisions.Remove(revision);
        _revisionExpandedBytes -= revision.ExpandedBytes;
    }

    private (BookProjectRevision Info, byte[] Bytes, long ExpandedBytes) RequireRevision(string revisionId) {
        if ((revisionId == null || string.IsNullOrWhiteSpace(revisionId))) throw new ArgumentException("A non-empty value is required.", nameof(revisionId));
        return _revisions.FirstOrDefault(item => item.Info.Id == revisionId) is var found && found.Info != null
            ? found : throw new KeyNotFoundException("The named revision does not exist in this project.");
    }
}
