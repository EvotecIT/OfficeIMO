namespace OfficeIMO.OpenDocument;

/// <summary>Common package lifecycle for ODT, ODS, and ODP documents.</summary>
public abstract partial class OdfDocument : IDisposable {
    private bool _disposed;
    private string? _sourcePath;

    internal OdfDocument(OdfPackage package, string? sourcePath) {
        Package = package ?? throw new ArgumentNullException(nameof(package));
        _sourcePath = sourcePath;
        Metadata = new OdfDocumentMetadata(this);
        Styles = new OdfStyleRepository(this);
    }

    internal OdfPackage Package { get; }

    /// <summary>Opens an ODT, ODS, or ODP package and returns its native document type.</summary>
    public static OdfDocument OpenAny(string path, OdfOpenOptions? options = null) {
        OdfPackage package = OdfPackage.Open(path, options, out string fullPath);
        return CreateForPackage(package, fullPath);
    }

    /// <summary>Opens an ODT, ODS, or ODP stream and returns its native document type.</summary>
    public static OdfDocument OpenAny(Stream stream, OdfOpenOptions? options = null) {
        return CreateForPackage(OdfPackage.Open(stream, options), null);
    }

    /// <summary>Document kind.</summary>
    public OdfDocumentKind Kind => Package.Kind;
    /// <summary>Source or current output version.</summary>
    public OdfVersion Version => Package.Version;
    /// <summary>Document metadata.</summary>
    public OdfDocumentMetadata Metadata { get; }
    /// <summary>Named and automatic styles stored in the document.</summary>
    public OdfStyleRepository Styles { get; }
    /// <summary>Non-fatal diagnostics produced while opening the package.</summary>
    public IReadOnlyList<OdfDiagnostic> Diagnostics => Package.Diagnostics;
    /// <summary>Most recent save-entry report.</summary>
    public OdfSaveReport? LastSaveReport { get; private set; }

    /// <summary>Saves to the original path used to open or first save the document.</summary>
    public void Save(OdfSaveOptions? options = null) {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(_sourcePath)) throw new InvalidOperationException("This document has no source path. Supply a destination path or stream.");
        Save(_sourcePath!, options);
    }

    /// <summary>Saves the document to a path using a same-directory temporary file.</summary>
    public void Save(string path, OdfSaveOptions? options = null) {
        ThrowIfDisposed();
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        string tempPath = Path.Combine(directory, "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            byte[] bytes = Render(options);
            File.WriteAllBytes(tempPath, bytes);
            ReplaceFile(tempPath, fullPath);
            _sourcePath = fullPath;
        } finally {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    /// <summary>Writes the document to a stream without closing it.</summary>
    public void Save(Stream destination, OdfSaveOptions? options = null) {
        ThrowIfDisposed();
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        byte[] bytes = Render(options);
        destination.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Asynchronously saves to a path.</summary>
    public async Task SaveAsync(string path, OdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        string tempPath = Path.Combine(directory, "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            byte[] bytes = Render(options);
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true)) {
#if NET8_0_OR_GREATER
                await stream.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
                await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
            }
            ReplaceFile(tempPath, fullPath);
            _sourcePath = fullPath;
        } finally {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    /// <summary>Asynchronously writes to a stream without closing it.</summary>
    public async Task SaveAsync(Stream destination, OdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        byte[] bytes = Render(options);
#if NET8_0_OR_GREATER
        await destination.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
        await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    /// <summary>Serializes the document to a byte array.</summary>
    public byte[] ToBytes(OdfSaveOptions? options = null) {
        ThrowIfDisposed();
        return Render(options);
    }

    /// <summary>Validates package and supported semantic invariants.</summary>
    public OdfValidationResult Validate() {
        ThrowIfDisposed();
        return OdfValidator.Validate(Package);
    }

    /// <summary>Inspects supported, preserved, and unsupported document features.</summary>
    public OdfFeatureReport InspectFeatures() {
        ThrowIfDisposed();
        return OdfFeatureInspector.Inspect(Package);
    }

    /// <summary>Releases document lifecycle state.</summary>
    public void Dispose() {
        _disposed = true;
    }

    internal XDocument GetXml(string partPath) {
        ThrowIfDisposed();
        return Package.GetXml(partPath);
    }

    internal void MarkPartDirty(string partPath) {
        ThrowIfDisposed();
        Package.MarkXmlDirty(partPath);
    }

    internal void AddDiagnostic(OdfDiagnostic diagnostic) {
        ThrowIfDisposed();
        Package.AddDiagnostic(diagnostic);
    }

    internal XElement GetBody(XName expectedBodyName) {
        XDocument content = GetXml("content.xml");
        XElement root = content.Root ?? throw new InvalidDataException("OpenDocument content has no root element.");
        XElement officeBody = root.Element(OdfNamespaces.Office + "body") ?? throw new InvalidDataException("OpenDocument content has no office:body.");
        return officeBody.Element(expectedBodyName) ?? throw new InvalidDataException($"OpenDocument body does not contain '{expectedBodyName}'.");
    }

    private byte[] Render(OdfSaveOptions? options) {
        byte[] bytes = Package.Write(options);
        LastSaveReport = Package.CreateSaveReport();
        return bytes;
    }

    private static OdfDocument CreateForPackage(OdfPackage package, string? sourcePath) {
        switch (package.Kind) {
            case OdfDocumentKind.Text: return new OdtDocument(package, sourcePath);
            case OdfDocumentKind.Spreadsheet: return new OdsDocument(package, sourcePath);
            case OdfDocumentKind.Presentation: return new OdpPresentation(package, sourcePath);
            default: throw new InvalidDataException("Unsupported OpenDocument package kind.");
        }
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(GetType().Name);
    }

    private static void ReplaceFile(string tempPath, string fullPath) {
        if (!File.Exists(fullPath)) {
            File.Move(tempPath, fullPath);
            return;
        }
        try {
            File.Replace(tempPath, fullPath, null);
        } catch (Exception ex) when (ex is PlatformNotSupportedException || ex is IOException) {
            File.Copy(tempPath, fullPath, overwrite: true);
            File.Delete(tempPath);
        }
    }
}
