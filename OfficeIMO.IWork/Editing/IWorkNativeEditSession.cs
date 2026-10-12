using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.IWork;

/// <summary>Edits a captured native source without changing the immutable source model.</summary>
/// <remarks>Instances are not thread safe. Supported mutations are qualified direct Keynote image translations;
/// other semantics and nested-index mutations are rejected. Unchanged saves support all modern read profiles.</remarks>
public sealed partial class IWorkNativeEditSession {
    private readonly IWorkSourceDocument _source;
    private readonly Dictionary<IWorkArchiveRecord, byte[]> _payloads = new();

    internal IWorkNativeEditSession(IWorkSourceDocument source) => _source = source;

    /// <summary>Gets the immutable source snapshot used by this edit session.</summary>
    public IWorkSourceDocument Source => _source;

    /// <summary>Encodes a native ZIP package, preserving all unmodified captured entry payloads.</summary>
    public byte[] SaveBytes(IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SaveBytesResult(options, cancellationToken).Bytes;

    /// <summary>Encodes a native ZIP package and returns its operation-level preservation report.</summary>
    public IWorkNativeSaveResult SaveBytesResult(IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        IWorkPreservationPackageWriter.Write(_source, _payloads, (options ?? new()).Snapshot(), cancellationToken);

    /// <summary>Writes a completed package at the caller stream's current position, leaves it open and returns the report.</summary>
    /// <remarks>Validation/encoding finish first. A copying failure can leave partial stream bytes.</remarks>
    public IWorkNativeSaveReport Save(Stream destination, IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        IWorkNativeSaveResult result = SaveBytesResult(options, cancellationToken);
        IWorkNativePackageWriter.Copy(result.Bytes, destination, cancellationToken); return result.Report;
    }

    /// <summary>Atomically publishes a native ZIP package. Existing destinations require explicit overwrite.</summary>
    /// <remarks>Directory-source members cannot be output targets. Failed encoding or staging preserves the existing destination.</remarks>
    public IWorkNativeSaveReport Save(string path, bool overwrite = false, IWorkNativeSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination path is required.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        RejectDirectoryMember(path);
        IWorkNativeSaveResult result = SaveBytesResult(options, cancellationToken);
        OfficeFileCommit.WriteAtomically(path, stream => IWorkNativePackageWriter.Copy(result.Bytes, stream, cancellationToken),
            cancellationToken, overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        return result.Report;
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
    public IWorkNativeEditSession BeginEdit() => new(this);
}
