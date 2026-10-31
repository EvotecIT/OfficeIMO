using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Loads a workbook with cancellation during native XLSB record projection.</summary>
    public static ExcelDocument Load(string path, ExcelLoadOptions? options, CancellationToken cancellationToken) {
        ExcelLoadOptions settings = (options ?? new ExcelLoadOptions()).Snapshot();
        settings.XlsbImportOptions.CancellationToken = cancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        ExcelDocument document = Load(path, settings);
        try { cancellationToken.ThrowIfCancellationRequested(); return document; }
        catch { document._persistenceMode = DocumentPersistenceMode.Explicit; document._copyPackageToSourceOnDispose = false; document._copyPackageToFilePathOnDispose = false; document.Dispose(); throw; }
    }

    /// <summary>Loads a caller-owned stream with cancellation during native XLSB record projection.</summary>
    public static ExcelDocument Load(Stream stream, ExcelLoadOptions? options, CancellationToken cancellationToken) {
        ExcelLoadOptions settings = (options ?? new ExcelLoadOptions()).Snapshot();
        settings.XlsbImportOptions.CancellationToken = cancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        ExcelDocument document = Load(stream, settings);
        try { cancellationToken.ThrowIfCancellationRequested(); return document; }
        catch { document._persistenceMode = DocumentPersistenceMode.Explicit; document._copyPackageToSourceOnDispose = false; document._copyPackageToFilePathOnDispose = false; document.Dispose(); throw; }
    }

    private string? _loadedPassword;

    private static ExcelDocument AssociateEncryptedWorkbook(ExcelDocument document, string? path, string password, Stream? source = null, ExcelLoadOptions? options = null) {
        try {
            if (source != null) OfficeIMO.Core.Internal.OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(source,
                options?.PersistenceMode ?? DocumentPersistenceMode.Explicit, nameof(source));
        } catch { document.Dispose(); throw; }
        document.FilePath = path;
        document._loadedPassword = password;
        document._persistenceMode = options?.PersistenceMode ?? DocumentPersistenceMode.Explicit;
        if (source != null) document._sourceStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(source, DocumentAccessMode.ReadWrite);
        return document;
    }

    private static ExcelDocument CreateFromUnifiedOptions(string? path, Stream? destination, ExcelCreateOptions options) {
        if (options.Template != null && options.TemplateStream != null)
            throw new ArgumentException("Choose a template path or a template stream.", nameof(options));
        if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && path == null && destination == null)
            throw new ArgumentException("SaveOnDispose requires an associated destination.", nameof(options));
        if (path != null && !options.Overwrite && File.Exists(path))
            throw new IOException($"The destination '{path}' already exists.");
        ExcelDocument document;
        var loadOptions = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit,
            OpenSettings = options.OpenSettings
        };
        if (options.Template != null) {
            using var source = File.OpenRead(options.Template);
            document = Load(source, loadOptions);
        } else {
            document = Load(options.TemplateStream ?? throw new ArgumentException("A template is required.", nameof(options)), loadOptions);
        }
        if (options.Theme != null) document.SetWorkbookThemeXml(CreateWorkbookThemeXml(options.Theme));
        document.FilePath = path;
        document._sourceStream = destination;
        document._persistenceMode = options.PersistenceMode;
        if (options.FirstSheetName != null && document.Sheets.Count == 0) document.AddSheet(options.FirstSheetName);
        return document;
    }

    private string? ResolveSavePassword(ExcelSaveOptions? options) =>
        options == null ? _loadedPassword : options.ResolvePassword(_loadedPassword);

    /// <summary>Encodes the workbook using one native-output options bag.</summary>
    public byte[] ToBytes(ExcelSaveOptions? options = null) {
        using var stream = new MemoryStream();
        Save(stream, options);
        return stream.ToArray();
    }

    private static ExcelFileFormat ResolvePathFormat(string path) =>
        ExcelDocumentLoadRouting.HasLegacyXlsExtension(path) ? ExcelFileFormat.Xls
        : ExcelDocumentLoadRouting.HasXlsbExtension(path) ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;

    private async Task SaveEncryptedFileAsync(string path, string password, ExcelSaveOptions? options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureDestinationFileWritable(path);
        EnsureDirectoryWritable(path);
        ExcelFileFormat format = ResolvePathFormat(path);
        if (format == ExcelFileFormat.Xlsx) AlignSpreadsheetDocumentTypeWithFilePath(path);
        using var stage = new MemoryStream();
        SaveEncrypted(stage, password, format, options);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeFileCommit.WriteAllBytesAsync(path, stage.ToArray(), cancellationToken: cancellationToken).ConfigureAwait(false);
        FilePath = path;
        _loadedPassword = password;
        DisableImplicitPersistenceAfterEncryptedSave();
    }
}
