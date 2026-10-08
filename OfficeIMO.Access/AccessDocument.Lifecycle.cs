using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access;

public sealed partial class AccessDocument {
    /// <summary>Creates an in-memory model without writing a file.</summary>
    public static AccessDocument Create(AccessCreateOptions? options = null) {
        options ??= new AccessCreateOptions(); options.Validate();
        return new AccessDocument(options.Format, DocumentAccessMode.ReadWrite, null);
    }
    /// <summary>Creates a model associated with a validated path; no directory or file is created.</summary>
    public static AccessDocument Create(string path, AccessCreateOptions? options = null) {
        var document = Create(options);
        try { document.ResolveTarget(path, new AccessSaveOptions { Format = document.Format }); document._path = Path.GetFullPath(path); return document; }
        catch { document.Dispose(); throw; }
    }
    /// <summary>Creates a model associated with a writable, seekable caller-owned stream. No bytes are written.</summary>
    public static AccessDocument Create(Stream stream, AccessCreateOptions? options = null) {
        OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
        var document = Create(options); document._destination = stream; return document;
    }
    /// <summary>Inspects a bounded native file without opening Access or executing active content.</summary>
    public static AccessInspection Inspect(string path, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate(); cancellationToken.ThrowIfCancellationRequested();
        using var stream = File.OpenRead(path);
        var inspection = InspectBytes(OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes), options, cancellationToken);
        ValidateSourcePath(path, inspection); return inspection;
    }
    /// <summary>Inspects caller-owned input. Seekable input starts at zero and its position is restored.</summary>
    public static AccessInspection Inspect(Stream stream, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate();
        return InspectBytes(OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes), options, cancellationToken);
    }
    private static AccessInspection InspectBytes(byte[] bytes, AccessLoadOptions options, CancellationToken cancellationToken) => AccessInspection.Read(bytes, options, cancellationToken);
    private static void ValidateSourcePath(string path, AccessInspection inspection) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension != ".mdb" && extension != ".accdb") throw new NotSupportedException("Native Access source paths require .mdb or .accdb. Use stream inspection for an independently identified payload.");
        if ((extension == ".mdb") != (inspection.Format == AccessFileFormat.Mdb)) throw new InvalidDataException("The Access source extension and detected file family disagree.");
    }
    private static AccessDocument FromBytes(byte[] bytes, AccessLoadOptions options, CancellationToken cancellationToken, string? path = null) {
        var inspection = InspectBytes(bytes, options, cancellationToken);
        if (path != null) ValidateSourcePath(path, inspection);
        var document = new AccessDocument(inspection.Format, options.AccessMode, inspection);
        document._inputLimit = options.MaxInputBytes; document._pageLimit = options.MaxPages;
        try {
            if (options.DecodeCatalog) {
                var native = new AccessNativeDatabase(document, bytes, options);
                document.NativeDatabase = native;
                if (native.CanDecode(out string reason)) native.LoadCatalog(cancellationToken);
                else document.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.catalog.unavailable", reason) });
            }
            return document;
        } catch { document.Dispose(); throw; }
    }
    /// <summary>Loads a bounded native catalog and selected schemas. User rows and large field values are decoded on demand.</summary>
    public static AccessDocument Load(string path, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate();
        using var stream = File.OpenRead(path);
        var document = FromBytes(OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes), options, cancellationToken, path); document._path = Path.GetFullPath(path); return document;
    }
    /// <summary>Loads caller-owned input into a bounded immutable snapshot without retaining or closing the stream.</summary>
    public static AccessDocument Load(Stream stream, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate(); return FromBytes(OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes), options, cancellationToken);
    }
    /// <summary>Asynchronously snapshots a bounded file and inspects its native header.</summary>
    public static async Task<AccessDocument> LoadAsync(string path, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate(); cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        var document = FromBytes(bytes, options, cancellationToken, path); document._path = Path.GetFullPath(path); return document;
    }
    /// <summary>Asynchronously snapshots caller-owned input, retaining its seekable position and leaving it open.</summary>
    public static async Task<AccessDocument> LoadAsync(Stream stream, AccessLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new AccessLoadOptions(); options.Validate();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        return FromBytes(bytes, options, cancellationToken);
    }
    /// <summary>Checks a loaded path against its full snapshot identity. Stream snapshots have no retained external source.</summary>
    public void ValidateSourceIdentity(CancellationToken cancellationToken = default) {
        EnsureNotDisposed(); cancellationToken.ThrowIfCancellationRequested();
        if (Inspection == null || _path == null) return;
        var current = Inspect(_path, new AccessLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, MaxInputBytes = _inputLimit, MaxPages = _pageLimit }, cancellationToken);
        if (current.Sha256 != Inspection.Sha256) throw new IOException("The Access source changed after loading. Reload before assessing native edits.");
    }
    private AccessFileFormat ResolveTarget(string? path, AccessSaveOptions options) {
        options.Validate();
        if (path == null) return options.Format ?? Format;
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An Access destination path is required.", nameof(path));
        string extension = Path.GetExtension(path);
        AccessFileFormat target = extension.ToLowerInvariant() switch {
            ".mdb" => AccessFileFormat.Mdb, ".accdb" => AccessFileFormat.Accdb,
            _ => throw new NotSupportedException("Access native destinations require .mdb or .accdb; compiled, add-in and packaged application formats are not qualified.")
        };
        if (options.Format.HasValue && target != options.Format.Value) throw new ArgumentException("The Access destination extension and explicit format disagree.", nameof(path));
        return target;
    }
    /// <summary>Assesses a native save without writing output or accepting unsupported codecs.</summary>
    public AccessOperationReport AssessSave(AccessSaveOptions? options = null, CancellationToken cancellationToken = default) => Assess(null, options, cancellationToken);
    /// <summary>Assesses native output for a path, including explicit format/extension agreement.</summary>
    public AccessOperationReport AssessSave(string path, AccessSaveOptions? options = null, CancellationToken cancellationToken = default) => Assess(path ?? throw new ArgumentNullException(nameof(path)), options, cancellationToken);
    private AccessOperationReport Assess(string? path, AccessSaveOptions? options, CancellationToken cancellationToken) {
        EnsureNotDisposed(); cancellationToken.ThrowIfCancellationRequested();
        if (HasActiveUpdate) throw new InvalidOperationException("Commit or roll back the Access update before assessing output.");
        options ??= new AccessSaveOptions(); AccessFileFormat target = ResolveTarget(path, options);
        AccessFormatProfile profile = options.Profile ?? (target == Format ? Profile : target == AccessFileFormat.Mdb ? AccessFormatProfile.Jet4 : AccessFormatProfile.Ace12);
        bool jet = profile == AccessFormatProfile.Jet3 || profile == AccessFormatProfile.Jet4;
        if (jet != (target == AccessFileFormat.Mdb)) throw new ArgumentException("The Access target profile and file family disagree.", nameof(options));
        var diagnostics = new List<AccessDiagnostic> { new AccessDiagnostic("access.native-write.unsupported", "Template-free native Access writing is not qualified. No output is produced.") };
        if (target != Format) diagnostics.Add(new AccessDiagnostic("access.conversion.unsupported", "MDB/ACCDB conversion and persistence codecs are unavailable. Feature-loss diagnostics do not enable output."));
        if (target != Format || profile != Profile) {
            if (Diagnostics.Any(x => x.Code.StartsWith("access.properties.", StringComparison.Ordinal)) || Properties.Values.Any(value => value is AccessOpaqueValue))
                diagnostics.Add(new AccessDiagnostic("access.conversion.loss.opaque-properties", "The target has no qualified mapping for this database's opaque property metadata; preservation cannot be assumed.", Id));
            foreach (var table in Tables) if (table.Diagnostics.Any(x => x.Code.StartsWith("access.properties.", StringComparison.Ordinal)) || table.Properties.Values.Any(value => value is AccessOpaqueValue) || table.Columns.Any(x => x.Properties.Values.Any(value => value is AccessOpaqueValue)))
                diagnostics.Add(new AccessDiagnostic("access.conversion.loss.opaque-properties", "The target has no qualified mapping for this table's opaque property metadata; preservation cannot be assumed.", table.Id));
        }
        foreach (var table in Tables) foreach (var column in table.Columns) {
            string? feature = target == AccessFileFormat.Mdb && column.DataType == AccessDataType.Complex ? "complex." + (column.ComplexDefinition?.Kind.ToString().ToLowerInvariant() ?? "unknown")
                : column.DataType == AccessDataType.Int64 && (jet || profile == AccessFormatProfile.Ace12 || profile == AccessFormatProfile.Ace14) ? "large-number"
                : column.DataType == AccessDataType.ExtendedDateTime && profile != AccessFormatProfile.Ace17 ? "extended-date"
                : target == AccessFileFormat.Mdb && column.IsRichText ? "rich-text" : null;
            if (column.IsCalculated && (jet || profile == AccessFormatProfile.Ace12)) feature = "calculated";
            if (feature != null) diagnostics.Add(new AccessDiagnostic("access.conversion.loss." + feature, "The requested target lacks this field's native feature. No implicit flattening, narrowing or precision loss is accepted.", column.Id));
        }
        if (CatalogStatus == AccessCatalogStatus.NotDecoded) diagnostics.Add(new AccessDiagnostic("access.catalog.not-decoded", "The source catalog and opaque application objects have not been decoded or qualified for preservation."));
        return new AccessOperationReport(Id, Revision, target, profile, diagnostics.AsReadOnly());
    }
    private void EnsureSaveAllowed() {
        EnsureNotDisposed();
        if (AccessMode == DocumentAccessMode.ReadOnly) throw new InvalidOperationException("This Access document is read-only and cannot be saved.");
    }
    /// <summary>Requests associated persistence. An unsupported writer fails before touching any destination.</summary>
    public void Save(AccessSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureSaveAllowed();
        if (_path != null) Save(_path, options, cancellationToken);
        else if (_destination != null) Save(_destination, options, cancellationToken);
        else throw new InvalidOperationException("Specify an Access destination path or stream.");
    }
    /// <summary>Requests native file output. Unsupported codecs fail before creating directories, files or temporary output.</summary>
    public void Save(string path, AccessSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureSaveAllowed(); AssessSave(path, options, cancellationToken).RequireNoLoss();
    }
    /// <summary>Requests native stream output. The caller's position, length and ownership remain unchanged on unsupported output.</summary>
    public void Save(Stream stream, AccessSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureSaveAllowed();
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        AssessSave(options, cancellationToken).RequireNoLoss();
    }
    /// <summary>Returns a failed asynchronous operation for an unsupported native writer without writing output.</summary>
    public Task SaveAsync(string path, AccessSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { Save(path, options, cancellationToken); return Task.CompletedTask; }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { return Task.FromCanceled(cancellationToken); }
        catch (Exception exception) { return Task.FromException(exception); }
    }
}
