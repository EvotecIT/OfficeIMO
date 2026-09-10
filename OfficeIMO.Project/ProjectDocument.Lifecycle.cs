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

    /// <summary>Creates a document associated with an XML, MPP, or MPT path; creation does not write a file.</summary>
    public static ProjectDocument Create(string path, DocumentCreateOptions? options = null) {
        var format = FormatFromPath(path);
        var document = Create();
        document._path = Path.GetFullPath(path);
        document._associatedFormat = format;
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

    /// <summary>Loads bounded XML or a qualified MPP14 document without calculating its schedule.</summary>
    public static ProjectDocument Load(string path, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An input path is required.", nameof(path));
        RejectGlobalTemplate(path);
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        using var stream = File.OpenRead(path);
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes);
        var document = ReadBytes(bytes, options, cancellationToken);
        document._path = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Loads a caller-owned stream. Seekable input is read from the start and its position is restored.</summary>
    public static ProjectDocument Load(Stream stream, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, options.PersistenceMode, nameof(stream));
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes);
        var document = ReadBytes(bytes, options, cancellationToken);
        document._associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, options.AccessMode);
        return document;
    }

    /// <summary>Asynchronously reads a file, then parses XML or qualified native input with cancellation and structural limits.</summary>
    public static async Task<ProjectDocument> LoadAsync(string path, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        RejectGlobalTemplate(path);
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        var document = ReadBytes(bytes, options, cancellationToken); document._path = Path.GetFullPath(path); return document;
    }

    /// <summary>Asynchronously reads caller-owned input while preserving seekable stream position.</summary>
    public static async Task<ProjectDocument> LoadAsync(Stream stream, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new ProjectLoadOptions(); options.ValidateLimits();
        OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, options.PersistenceMode, nameof(stream));
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        var document = ReadBytes(bytes, options, cancellationToken);
        document._associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, options.AccessMode); return document;
    }

    /// <summary>Assesses the current model for the selected format without committing output or calculating a schedule.</summary>
    public ProjectReport AssessSave(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed(); options ??= new ProjectSaveOptions(); options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        return AssessFormat(options, ResolveFormat(options, options.Format == ProjectFileFormat.Automatic ? _path : null), cancellationToken);
    }

    /// <summary>Assesses a destination's format, including conversion loss, against the current model revision.</summary>
    public ProjectReport AssessSave(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed(); options ??= new ProjectSaveOptions(); return AssessFormat(options, ResolveFormat(options, path), cancellationToken);
    }

    /// <summary>Serializes to XML text; no destination is rebound and the document remains modified.</summary>
    public string ToXml(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        options ??= new ProjectSaveOptions();
        options.Validate();
        var textOptions = WithFormat(options, ProjectFileFormat.Xml, false);
        return Encoding.UTF8.GetString(Serialize(textOptions, cancellationToken).Bytes);
    }

    /// <summary>Saves to the associated file or seekable stream.</summary>
    public void Save(ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (_path != null) Save(_path, options, cancellationToken);
        else if (_associatedStream != null) Save(_associatedStream, options, cancellationToken);
        else throw new InvalidOperationException("No associated destination. Specify a Project path or stream.");
    }

    /// <summary>Atomically saves to the format selected by the extension and associates the document with that path after success.</summary>
    public void Save(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions(); options = WithFormat(options, ResolveFormat(options, path));
        var prepared = Serialize(options, cancellationToken); byte[] bytes = prepared.Bytes;
        OfficeFileCommit.Write(path, output => {
            cancellationToken.ThrowIfCancellationRequested(); output.Write(bytes, 0, bytes.Length); cancellationToken.ThrowIfCancellationRequested();
        }, options.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        _path = Path.GetFullPath(path); _associatedStream = null; AcceptSaved(prepared);
    }

    /// <summary>Saves to a caller-owned stream. Seekable streams are replaced; non-seekable streams receive bytes at their current position.</summary>
    /// <remarks>Stream writes cannot promise atomic rollback on I/O failure. File saves use atomic replacement.</remarks>
    public void Save(Stream stream, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions();
        options = WithFormat(options, ResolveFormat(options));
        var prepared = Serialize(options, cancellationToken); byte[] bytes = prepared.Bytes;
        cancellationToken.ThrowIfCancellationRequested(); OfficeStreamWriter.WriteAllBytes(stream, bytes);
        _associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, _accessMode); _path = null; AcceptSaved(prepared);
    }

    /// <summary>Asynchronously commits a complete Project file after validation and serialization.</summary>
    public async Task SaveAsync(string path, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions(); options = WithFormat(options, ResolveFormat(options, path));
        var prepared = Serialize(options, cancellationToken); byte[] bytes = prepared.Bytes;
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes,
            options.FileConflictPolicy == OfficeConversionFileConflictPolicy.Replace ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists, cancellationToken).ConfigureAwait(false);
        _path = Path.GetFullPath(path); _associatedStream = null; AcceptSaved(prepared);
    }

    /// <summary>Asynchronously writes the selected format to a caller-owned stream; failure may leave a partial destination.</summary>
    public async Task SaveAsync(Stream stream, ProjectSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureMutable(); options ??= new ProjectSaveOptions();
        options = WithFormat(options, ResolveFormat(options));
        var prepared = Serialize(options, cancellationToken); byte[] bytes = prepared.Bytes;
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        _associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, _accessMode); _path = null; AcceptSaved(prepared);
    }

    private ProjectSerialization Serialize(ProjectSaveOptions options, CancellationToken token) {
        EnsureNotDisposed(); options.Validate();
        if (_batchDepth != 0) throw new InvalidOperationException("Complete the update scope before saving.");
        long revision = Revision; var format = ResolveFormat(options);
        var assessment = AssessFormat(options, format, token, includeNativePlan: false); assessment.ThrowIfErrors();
        if (options.LossPolicy == OfficeConversionLossPolicy.Block) assessment.RequireNoLoss();
        byte[] bytes; ProjectNativeSource? native = null; ProjectMpxSource? mpx = null;
        if (format == ProjectFileFormat.Mpx4) {
            var plan = ProjectMpxWriter.Plan(this, WithFormat(options, format), true, token);
            plan.Report.ThrowIfErrors(); if (options.LossPolicy == OfficeConversionLossPolicy.Block) plan.Report.RequireNoLoss();
            bytes = plan.Bytes!;
            if (ProjectMpxWriter.CanRetain(this, options)) mpx = MpxSource;
            else {
                var snapshot = ProjectModelSnapshot.Capture(this, token);
                mpx = new ProjectMpxSource { Bytes = bytes, ModelRevision = revision,
                    CodePage = (int?)options.MpxEncoding ?? MpxSource?.CodePage ?? 1252, Separator = options.MpxSeparator ?? MpxSource?.Separator ?? ',',
                    CurrencyPosition = MpxSource?.CurrencyPosition ?? "1", DateFormat = MpxSource?.DateFormat ?? "0", BarDateFormat = MpxSource?.BarDateFormat ?? "0", Comments = MpxSource?.Comments ?? Array.Empty<string[]>(),
                    Unrepresented = RetainedModelLosses(plan.Report, snapshot)
                };
            }
        } else if (IsNativeFormat(format)) {
            token.ThrowIfCancellationRequested();
            if (CanRetainNative(options, format)) {
                if (NativeSource!.Bytes.Length > options.MaxOutputBytes) throw new InvalidDataException("Native output exceeds the configured byte limit.");
                bytes = NativeSource.Bytes; native = NativeSource;
            } else {
                var plan = ProjectNativeWriter.Plan(this, WithFormat(options, format), true, token);
                plan.Report.ThrowIfErrors(); if (options.LossPolicy == OfficeConversionLossPolicy.Block) plan.Report.RequireNoLoss();
                bytes = plan.Bytes!;
                if (!OfficeCompoundFileReader.TryRead(bytes, new OfficeCompoundReadOptions(int.MaxValue, int.MaxValue, bytes.Length, bytes.Length), token, out var file, out var error) || file == null)
                    throw new InvalidDataException("Prepared native output could not be read: " + error);
                var preparedProfile = ProjectNativeProfile.Detect(file);
                var preparedHeader = ProjectNativeProperties.Read(file.Streams[preparedProfile.Header], token);
                var snapshot = ProjectModelSnapshot.Capture(this, token);
                native = new ProjectNativeSource(bytes, new ProjectNativeInfo(preparedHeader.TryGetValue(0x35400010, out var producer) ? producer.Unicode() : null, file)) {
                    Snapshot = snapshot, ModelRevision = revision,
                    UnrepresentedValues = RetainedModelLosses(plan.Report, snapshot)
                };
            }
        } else bytes = ProjectXmlCodec.Write(this, WithFormat(options, format, NativeSource != null || MpxSource != null ? false : (bool?)null), token);
        token.ThrowIfCancellationRequested();
        if (Revision != revision) throw new InvalidOperationException("The project changed while serialization was in progress.");
        return new ProjectSerialization(bytes, format, revision, native, mpx);
    }
    private static IReadOnlyList<ProjectDiagnostic> RetainedModelLosses(ProjectReport report, Dictionary<string, object?> snapshot) =>
        report.Diagnostics.Where(d => d.RepresentsLoss && d.Location != "/" &&
            (snapshot.ContainsKey(d.Location) || snapshot.Keys.Any(k => k.StartsWith(d.Location + "/", StringComparison.Ordinal) || k.StartsWith(d.Location + "[", StringComparison.Ordinal)))).ToArray();
    private void AcceptSaved(ProjectSerialization prepared) {
        LastSavedBytes = prepared.Bytes; _savedRevision = prepared.Revision; _associatedFormat = prepared.Format;
        if (prepared.Native != null) NativeSource = prepared.Native;
        if (prepared.Mpx != null) MpxSource = prepared.Mpx;
    }
    private sealed class ProjectSerialization {
        internal readonly byte[] Bytes;
        internal readonly ProjectFileFormat Format;
        internal readonly long Revision;
        internal readonly ProjectNativeSource? Native;
        internal readonly ProjectMpxSource? Mpx;
        internal ProjectSerialization(byte[] bytes, ProjectFileFormat format, long revision, ProjectNativeSource? native, ProjectMpxSource? mpx) {
            Bytes = bytes; Format = format; Revision = revision; Native = native; Mpx = mpx;
        }
    }
    private static ProjectDocument ReadBytes(byte[] bytes, ProjectLoadOptions options, CancellationToken token) =>
        ProjectNativeCodec.IsCompound(bytes) ? ProjectNativeCodec.Read(bytes, options, token) :
        ProjectMpxRecords.IsMpx(bytes) ? ProjectMpxCodec.Read(bytes, options, token) : ProjectXmlCodec.Read(bytes, options, token);
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
        Source = null; NativeSource = null; MpxSource = null; LastSavedBytes = null;
        _disposed = true;
    }
}
