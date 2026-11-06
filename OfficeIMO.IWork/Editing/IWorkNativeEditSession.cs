using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.IWork;

/// <summary>Edits a captured native source without changing the immutable source model.</summary>
/// <remarks>Instances are not thread safe. Supported mutations are direct Keynote image moves/resizes and existing
/// Numbers numeric/boolean cells and fixed-length plain Pages body or direct Keynote text-box paragraphs in bounded modern profiles. Nested-index mutations are rejected.
/// Unchanged saves support all modern read profiles.</remarks>
public sealed partial class IWorkNativeEditSession {
    private readonly IWorkSourceDocument _source;
    private readonly Dictionary<IWorkArchiveRecord, byte[]> _payloads = new();

    internal IWorkNativeEditSession(IWorkSourceDocument source) => _source = source;

    /// <summary>Gets the immutable source snapshot used by this edit session.</summary>
    public IWorkSourceDocument Source => _source;

    /// <summary>Encodes a native ZIP package, preserving all unmodified captured entry payloads.</summary>
    public byte[] ToBytes(IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        ToBytes(out _, options, cancellationToken);

    /// <summary>Encodes a native ZIP package and returns its actual operation-level preservation evidence.</summary>
    public byte[] ToBytes(out IWorkNativeSaveReport report, IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) {
        var result = IWorkPreservationPackageWriter.Write(_source, _payloads, (options ?? new()).Snapshot(), cancellationToken);
        report = result.Report;
        return result.Bytes;
    }

    /// <summary>Writes a completed package at the caller stream's current position, leaves it open and returns the report.</summary>
    /// <remarks>Validation/encoding finish first. A copying failure can leave partial stream bytes.</remarks>
    public IWorkNativeSaveReport Save(Stream destination, IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        byte[] bytes = ToBytes(out var report, options, cancellationToken);
        IWorkNativePackageWriter.Copy(bytes, destination, cancellationToken); return report;
    }

    /// <summary>Atomically publishes a native ZIP package. Existing destinations require explicit overwrite.</summary>
    /// <remarks>Directory-source members cannot be output targets. Failed encoding or staging preserves the existing destination.</remarks>
    public IWorkNativeSaveReport Save(string path, bool overwrite = false, IWorkNativeSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        cancellationToken.ThrowIfCancellationRequested();
        RejectDirectoryMember(path);
        byte[] bytes = ToBytes(out var report, options, cancellationToken);
        OfficeFileCommit.WriteAtomically(path, stream => IWorkNativePackageWriter.Copy(bytes, stream, cancellationToken),
            cancellationToken, overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        FilePath = Path.GetFullPath(path);
        return report;
    }

    private void RejectDirectoryMember(string path) {
        if (_source.SourceDirectoryPath == null) return;
        string full = Path.GetFullPath(path), parent = Path.GetDirectoryName(full)!;
        string target = Path.Combine(OfficePathIdentity.ResolvePhysicalPath(parent), Path.GetFileName(full));
        if (OfficePathIdentity.IsSameOrDescendant(target, _source.SourceDirectoryPath))
            throw new InvalidOperationException("A native output cannot be published inside its captured directory source.");
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Starts a separate preservation-aware native edit session over this captured snapshot.</summary>
    public IWorkNativeEditSession BeginEdit() {
        if (LegacyXml != null)
            throw new NotSupportedException("Native edit and save are unavailable for iWork '09 XML packages.");
        return new(this);
    }
}
