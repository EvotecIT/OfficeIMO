using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private string? _filePath;

    /// <summary>Gets the associated path, or null for a document created or loaded in memory.</summary>
    public string? FilePath => _filePath;

    /// <summary>Creates an empty PDF associated with a path. The file is written on Save.</summary>
    public static PdfDocument Create(string filePath, PdfCreateOptions? options = null) =>
        new PdfDocument(options) { _filePath = ValidateOutputPath(filePath) };

    /// <summary>Saves to the associated path, rejecting serialization loss by default.</summary>
    public void Save() => Save(_filePath ?? throw new InvalidOperationException("The document has no associated path."));

    /// <summary>Saves to the associated path using explicit native-output options.</summary>
    public void Save(PdfSaveOptions options) =>
        Save(_filePath ?? throw new InvalidOperationException("The document has no associated path."), options);

    /// <summary>Serializes PDF bytes and checks any new native-output fidelity findings.</summary>
    public byte[] ToBytes(PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_source != null) return SerializeBytes(cancellationToken);
        using var scope = new PdfOutputDiagnosticScope();
        byte[] bytes = SerializeBytes(cancellationToken);
        OfficeConversionReport report = PdfConversionEvidence.CreateReport("Pdf", new[] { scope.Report });
        if (options?.AllowLoss != true) report.ThrowIfLoss();
        return bytes;
    }

    /// <summary>Atomically saves the native document to a file and associates that path.</summary>
    public void Save(string path, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string fullPath = ValidateOutputPath(path);
        ThrowIfTextEncodingPreflightFails(cancellationToken);
        OfficeFileCommit.Write(fullPath, output => WriteNativeWithLossCheck(output, options, cancellationToken),
            cancellationToken, ResolveConflictPolicy(options));
        _filePath = fullPath;
    }

    /// <summary>Writes native PDF to a caller-owned stream after validating fidelity.</summary>
    public void Save(Stream stream, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        SaveNativeStream(stream, options, cancellationToken);
    }

    /// <summary>Asynchronously saves the native document to its associated path.</summary>
    public Task SaveAsync(CancellationToken cancellationToken = default) =>
        SaveAsync(_filePath ?? throw new InvalidOperationException("The document has no associated path."), null, cancellationToken);

    /// <summary>Asynchronously saves the native document to a file and associates that path.</summary>
    public async Task SaveAsync(string path, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string fullPath = ValidateOutputPath(path);
        ThrowIfTextEncodingPreflightFails(cancellationToken);
        await OfficeFileCommit.WriteAsync(fullPath, output => WriteNativeWithLossCheck(output, options, cancellationToken),
            ResolveConflictPolicy(options), cancellationToken).ConfigureAwait(false);
        _filePath = fullPath;
    }

    /// <summary>Asynchronously writes native PDF to a caller-owned stream after validating fidelity.</summary>
    public async Task SaveAsync(Stream stream, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        await SaveNativeStreamAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    private static OfficeFileCommit.ConflictPolicy ResolveConflictPolicy(PdfSaveOptions? options) =>
        (options?.ConflictPolicy ?? OfficeConversionFileConflictPolicy.Replace) switch {
            OfficeConversionFileConflictPolicy.Replace => OfficeFileCommit.ConflictPolicy.Replace,
            OfficeConversionFileConflictPolicy.FailIfExists => OfficeFileCommit.ConflictPolicy.FailIfExists,
            _ => throw new ArgumentOutOfRangeException(nameof(options))
        };

    /// <summary>Completes this in-memory document's lifetime. Caller-owned streams remain open.</summary>
    public void Dispose() { }

    /// <summary>Completes this in-memory document's lifetime asynchronously.</summary>
    public Task DisposeAsync() { Dispose(); return Task.CompletedTask; }
}
