using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Creates an in-memory document with deterministic working-time defaults, USD currency, and no implicit dates or calendar.</summary>
    public static ProjectDocument Create() {
        var document = new ProjectDocument { _savedRevision = -1 };
        document.Settings.CurrencyCode = "USD";
        return document;
    }

    /// <summary>Creates a document associated with an XML path; creation does not write a file.</summary>
    public static ProjectDocument Create(string path, DocumentCreateOptions? options = null) {
        ValidateXmlPath(path);
        var document = Create();
        document._path = Path.GetFullPath(path);
        document.SetCreateOptions(options);
        return document;
    }

    /// <summary>Creates a document associated with a caller-owned writable seekable stream.</summary>
    public static ProjectDocument Create(Stream stream, DocumentCreateOptions? options = null) {
        OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
        var document = Create(); document._associatedStream = stream; document.SetCreateOptions(options); return document;
    }

    private void SetCreateOptions(DocumentCreateOptions? options) {
        _persistenceMode = options?.PersistenceMode ?? DocumentPersistenceMode.Explicit;
        if (!Enum.IsDefined(typeof(DocumentPersistenceMode), _persistenceMode)) throw new ArgumentOutOfRangeException(nameof(options));
    }

    /// <summary>Parses a Project XML string with bounded, DTD-disabled parsing.</summary>
    public static ProjectDocument Parse(string xml, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) throw new ArgumentException("Text input has no associated destination for SaveOnDispose.", nameof(options));
        if (xml.Length > options.MaxCharacters || Encoding.UTF8.GetByteCount(xml) > options.MaxInputBytes) throw new InvalidDataException("Project text exceeds the input limits.");
        // A .NET string has already been decoded. Remove any encoding declaration by reading
        // through TextReader rather than treating a UTF-16 declaration as UTF-8 byte input.
        using var text = new StringReader(xml);
        var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaxCharacters };
        using var raw = System.Xml.XmlReader.Create(text, settings);
        using var reader = new OfficeXmlLimitingReader(raw, "MSPDI", options.MaxDepth, options.MaxElements, options.MaxAttributes, cancellationToken);
        var parsed = System.Xml.Linq.XDocument.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        parsed.Declaration = new System.Xml.Linq.XDeclaration("1.0", "utf-8", parsed.Declaration?.Standalone);
        byte[] bytes;
        using (var buffer = new OfficeBoundedMemoryStream(options.MaxInputBytes)) {
            using (var writer = System.Xml.XmlWriter.Create(buffer, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, NewLineHandling = System.Xml.NewLineHandling.Entitize })) parsed.Save(writer);
            bytes = buffer.ToArray();
        }
        return ProjectXmlCodec.Read(bytes, options, cancellationToken);
    }

    /// <summary>Loads XML from a file. Native binary formats are detected and explicitly rejected by this codec.</summary>
    public static ProjectDocument Load(string path, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An input path is required.", nameof(path));
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        using var stream = File.OpenRead(path);
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes);
        var document = ProjectXmlCodec.Read(bytes, options, cancellationToken);
        document._path = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Loads a caller-owned stream. Seekable input is read from the start and its position is restored.</summary>
    public static ProjectDocument Load(Stream stream, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, options.PersistenceMode, nameof(stream));
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes);
        var document = ProjectXmlCodec.Read(bytes, options, cancellationToken);
        document._associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, options.AccessMode);
        return document;
    }

    /// <summary>Asynchronously reads a file, then parses its XML with cancellation and structural limits.</summary>
    public static async Task<ProjectDocument> LoadAsync(string path, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        var document = ProjectXmlCodec.Read(bytes, options, cancellationToken); document._path = Path.GetFullPath(path); return document;
    }

    /// <summary>Asynchronously reads caller-owned input while preserving seekable stream position.</summary>
    public static async Task<ProjectDocument> LoadAsync(Stream stream, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, options.PersistenceMode, nameof(stream));
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        var document = ProjectXmlCodec.Read(bytes, options, cancellationToken);
        document._associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, options.AccessMode); return document;
    }

    /// <summary>Assesses the current model and XML preservation boundary without writing or calculating.</summary>
    public ProjectReport AssessSave(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed(); options ??= new ProjectSaveOptions(); options.Validate();
        return Validate(cancellationToken);
    }

    /// <summary>Assesses an XML destination. Changing the extension cannot enable a native MPP writer.</summary>
    public ProjectReport AssessSave(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        ValidateXmlPath(path); return AssessSave(options, cancellationToken);
    }

    /// <summary>Serializes to XML text; no destination is rebound and the document remains modified.</summary>
    public string ToXml(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectSaveOptions();
        var textOptions = new ProjectSaveOptions { Indent = options.Indent, LossPolicy = options.LossPolicy, FileConflictPolicy = options.FileConflictPolicy, MaxOutputBytes = options.MaxOutputBytes, PreserveUnchangedBytes = false };
        return Encoding.UTF8.GetString(Serialize(textOptions, cancellationToken));
    }

    /// <summary>Saves to the associated file or seekable stream.</summary>
    public void Save(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (_path != null) Save(_path, options, cancellationToken);
        else if (_associatedStream != null) Save(_associatedStream, options, cancellationToken);
        else throw new InvalidOperationException("No associated destination. Specify an XML path or stream.");
    }

    /// <summary>Atomically saves to an XML file and associates the document with that path after success.</summary>
    public void Save(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); ValidateXmlPath(path); options ??= new ProjectSaveOptions();
        byte[] bytes = Serialize(options, cancellationToken);
        OfficeFileCommit.Write(path, output => {
            cancellationToken.ThrowIfCancellationRequested(); output.Write(bytes, 0, bytes.Length); cancellationToken.ThrowIfCancellationRequested();
        }, options.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        _path = Path.GetFullPath(path); _associatedStream = null; AcceptSaved(bytes);
    }

    /// <summary>Saves to a caller-owned stream. Seekable streams are replaced; non-seekable streams receive bytes at their current position.</summary>
    /// <remarks>Stream writes cannot promise atomic rollback on I/O failure. File saves use atomic replacement.</remarks>
    public void Save(Stream stream, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions();
        byte[] bytes = Serialize(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested(); OfficeStreamWriter.WriteAllBytes(stream, bytes);
        _associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, _accessMode); _path = null; AcceptSaved(bytes);
    }

    /// <summary>Asynchronously commits a complete XML file after validation and serialization.</summary>
    public async Task SaveAsync(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); ValidateXmlPath(path); options ??= new ProjectSaveOptions();
        byte[] bytes = Serialize(options, cancellationToken);
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes,
            options.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists, cancellationToken).ConfigureAwait(false);
        _path = Path.GetFullPath(path); _associatedStream = null; AcceptSaved(bytes);
    }

    /// <summary>Asynchronously writes XML to a caller-owned stream; failure may leave a partial destination.</summary>
    public async Task SaveAsync(Stream stream, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions();
        byte[] bytes = Serialize(options, cancellationToken);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        _associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, _accessMode); _path = null; AcceptSaved(bytes);
    }

    private byte[] Serialize(ProjectSaveOptions options, CancellationToken token) {
        EnsureNotDisposed(); options.Validate();
        if (_batchDepth != 0) throw new InvalidOperationException("Complete the update scope before saving.");
        var assessment = AssessSave(options, token); assessment.ThrowIfErrors();
        if (options.LossPolicy == OfficeConversionLossPolicy.Block) assessment.RequireNoLoss();
        return ProjectXmlCodec.Write(this, options, token);
    }
    private void AcceptSaved(byte[] bytes) { LastSavedBytes = bytes; _savedRevision = Revision; }
    private static void ValidateXmlPath(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An XML path is required.", nameof(path));
        if (!string.Equals(Path.GetExtension(path), ".xml", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Only .xml output is supported. MPP, MPT, and MPX are separate native format capabilities.");
    }
    internal static ProjectDocument CreateForRead() => new ProjectDocument { Loading = true };
    internal void FinishRead(ProjectLoadOptions options) {
        OfficeDocumentLifecycle.Validate(options.AccessMode, options.PersistenceMode, "Project document");
        _accessMode = options.AccessMode; _persistenceMode = options.PersistenceMode; Loading = false; _savedRevision = Revision;
    }
    internal void DisposeFailedRead() { Loading = false; _persistenceMode = DocumentPersistenceMode.Explicit; _disposed = true; }

    /// <summary>Disposes the document. Caller-owned streams remain open. Explicit persistence is the default.</summary>
    public void Dispose() {
        if (_disposed) return;
        if (_persistenceMode == DocumentPersistenceMode.SaveOnDispose && IsModified) Save();
        Source = null; LastSavedBytes = null;
        _disposed = true;
    }
}
