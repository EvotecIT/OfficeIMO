using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.OpenDocument;

/// <summary>Common package lifecycle for ODT, ODS, and ODP documents.</summary>
public abstract partial class OdfDocument {
    private string? _sourcePath;

    internal OdfDocument(OdfPackage package, string? sourcePath) {
        Package = package ?? throw new ArgumentNullException(nameof(package));
        _sourcePath = sourcePath;
        Metadata = new OdfDocumentMetadata(this);
        Styles = new OdfStyleRepository(this);
    }

    internal OdfPackage Package { get; }

    /// <summary>Loads an ODT, ODS, or ODP package and returns its native document type.</summary>
    public static OdfDocument Load(string path, OdfLoadOptions? options = null) {
        OdfPackage package = OdfPackage.Load(path, options, out string fullPath);
        return CreateForPackage(package, fullPath);
    }

    /// <summary>Loads an ODT, ODS, or ODP stream and returns its native document type.</summary>
    public static OdfDocument Load(Stream stream, OdfLoadOptions? options = null) {
        return CreateForPackage(OdfPackage.Load(stream, options), null);
    }

    /// <summary>Asynchronously loads an ODT, ODS, or ODP package from a path.</summary>
    public static async Task<OdfDocument> LoadAsync(
        string path,
        OdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        EnsurePath(path);
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException($"OpenDocument file '{fullPath}' does not exist.", fullPath);
        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            81920,
            useAsync: true);
        OdfPackage package = await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false);
        return CreateForPackage(package, fullPath);
    }

    /// <summary>Asynchronously loads an ODT, ODS, or ODP package from a caller-owned stream.</summary>
    public static async Task<OdfDocument> LoadAsync(
        Stream stream,
        OdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        OdfPackage package = await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false);
        return CreateForPackage(package, null);
    }

    /// <summary>Document kind.</summary>
    public OdfDocumentKind Kind => Package.Kind;
    /// <summary>Gets the file path associated with the document, if any.</summary>
    public string? FilePath => _sourcePath;
    /// <summary>Source or current output version.</summary>
    public OdfVersion Version => Package.Version;
    /// <summary>Document metadata.</summary>
    public OdfDocumentMetadata Metadata { get; }
    /// <summary>Named and automatic styles stored in the document.</summary>
    public OdfStyleRepository Styles { get; }
    /// <summary>Non-fatal diagnostics produced while opening the package.</summary>
    public IReadOnlyList<OdfDiagnostic> Diagnostics => Package.Diagnostics;
    /// <summary>Saves to the current path and returns the serialized bytes with entry-level diagnostics.</summary>
    public OdfSaveResult Save(OdfSaveOptions? options = null) {
        if (string.IsNullOrEmpty(_sourcePath)) throw new InvalidOperationException("This document has no source path. Supply a destination path or stream.");
        return Save(_sourcePath!, options);
    }

    /// <summary>Saves to a path and returns the serialized bytes with entry-level diagnostics.</summary>
    public OdfSaveResult Save(string path, OdfSaveOptions? options = null) {
        EnsurePath(path);
        string fullPath = Path.GetFullPath(path);
        byte[] bytes = Render(options, out OdfSaveReport report);
        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
        _sourcePath = fullPath;
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Saves an independent copy and returns entry-level diagnostics without changing the associated path.</summary>
    public OdfSaveResult SaveCopy(string path, OdfSaveOptions? options = null) {
        EnsurePath(path);
        string fullPath = Path.GetFullPath(path);
        byte[] bytes = Render(options, out OdfSaveReport report);
        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Writes to a stream and returns the serialized bytes with entry-level diagnostics.</summary>
    public OdfSaveResult Save(Stream destination, OdfSaveOptions? options = null) {
        byte[] bytes = Render(options, out OdfSaveReport report);
        OfficeStreamWriter.WriteAllBytes(destination, bytes);
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Asynchronously saves to the original path used to load or first save the document.</summary>
    public Task<OdfSaveResult> SaveAsync(CancellationToken cancellationToken = default) =>
        SaveAsync(options: null, cancellationToken);

    /// <summary>Asynchronously saves to the original path with optional save settings.</summary>
    public Task<OdfSaveResult> SaveAsync(OdfSaveOptions? options, CancellationToken cancellationToken = default) {
        if (string.IsNullOrEmpty(_sourcePath)) {
            throw new InvalidOperationException("This document has no source path. Supply a destination path or stream.");
        }
        return SaveAsync(_sourcePath!, options, cancellationToken);
    }

    /// <summary>Asynchronously saves to a path and returns the serialized bytes with entry-level diagnostics.</summary>
    public async Task<OdfSaveResult> SaveAsync(string path, OdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsurePath(path);
        string fullPath = Path.GetFullPath(path);
        byte[] bytes = Render(options, out OdfSaveReport report);
        await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
        _sourcePath = fullPath;
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Asynchronously saves an independent copy and returns entry-level diagnostics without changing the associated path.</summary>
    public async Task<OdfSaveResult> SaveCopyAsync(string path, OdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsurePath(path);
        string fullPath = Path.GetFullPath(path);
        byte[] bytes = Render(options, out OdfSaveReport report);
        await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Asynchronously writes to a stream and returns the serialized bytes with entry-level diagnostics.</summary>
    public async Task<OdfSaveResult> SaveAsync(Stream destination, OdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        byte[] bytes = Render(options, out OdfSaveReport report);
        await OfficeStreamWriter.WriteAllBytesAsync(destination, bytes, cancellationToken).ConfigureAwait(false);
        CompleteSave();
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Serializes the document to a byte array.</summary>
    public byte[] ToBytes(OdfSaveOptions? options = null) {
        return Serialize(options).RequireValue();
    }

    /// <summary>Serializes the document in a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(OdfSaveOptions? options = null) => new MemoryStream(ToBytes(options));

    private static void EnsurePath(string path) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }
    }

    /// <summary>
    /// Serializes without accepting the current dirty state and returns the bytes with entry-level diagnostics.
    /// </summary>
    public OdfSaveResult Serialize(OdfSaveOptions? options = null) {
        byte[] bytes = Render(options, out OdfSaveReport report);
        return new OdfSaveResult(bytes, report);
    }

    /// <summary>Validates package and supported semantic invariants.</summary>
    public OdfValidationResult Validate() {
        return OdfValidator.Validate(Package);
    }

    /// <summary>Inspects supported, preserved, and unsupported document features.</summary>
    public OdfFeatureReport InspectFeatures() {
        return OdfFeatureInspector.Inspect(Package);
    }

    internal XDocument GetXml(string partPath) {
        return Package.GetXml(partPath);
    }

    internal void MarkPartDirty(string partPath) {
        Package.MarkXmlDirty(partPath);
    }

    internal void AddDiagnostic(OdfDiagnostic diagnostic) {
        Package.AddDiagnostic(diagnostic);
    }

    internal XElement GetBody(XName expectedBodyName) {
        XDocument content = GetXml("content.xml");
        XElement root = content.Root ?? throw new InvalidDataException("OpenDocument content has no root element.");
        XElement officeBody = root.Element(OdfNamespaces.Office + "body") ?? throw new InvalidDataException("OpenDocument content has no office:body.");
        return officeBody.Element(expectedBodyName) ?? throw new InvalidDataException($"OpenDocument body does not contain '{expectedBodyName}'.");
    }

    private byte[] Render(OdfSaveOptions? options, out OdfSaveReport report) {
        byte[] bytes = Package.Write(options);
        report = Package.CreateSaveReport();
        return bytes;
    }

    private void CompleteSave() {
        Package.AcceptChanges();
    }

    private static OdfDocument CreateForPackage(OdfPackage package, string? sourcePath) {
        switch (package.Kind) {
            case OdfDocumentKind.Text: return new OdtDocument(package, sourcePath);
            case OdfDocumentKind.Spreadsheet: return new OdsDocument(package, sourcePath);
            case OdfDocumentKind.Presentation: return new OdpPresentation(package, sourcePath);
            default: throw new InvalidDataException("Unsupported OpenDocument package kind.");
        }
    }

    internal static async Task<OdfPackage> LoadPackageAsync(
        Stream stream,
        OdfLoadOptions? options,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OpenDocument stream must be readable.", nameof(stream));
        OdfLoadOptions effective = (options ?? new OdfLoadOptions()).Normalize();
        try {
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(
                stream,
                cancellationToken,
                effective.MaxPackageBytes).ConfigureAwait(false);
            return OdfPackage.Load(bytes, effective);
        } catch (InvalidDataException ex) when (ex.Message.IndexOf("configured maximum size", StringComparison.Ordinal) >= 0) {
            throw new InvalidDataException(
                $"OpenDocument stream exceeds MaxPackageBytes ({effective.MaxPackageBytes}).",
                ex);
        }
    }

}
