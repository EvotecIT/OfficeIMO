using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Write;
using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>
    /// Encrypts a workbook in an explicitly selected physical format. XLS uses classic BIFF8 RC4 for legacy compatibility;
    /// XLSX and XLSB use the existing Office package encryption profile. XLS does not provide modern encryption strength.
    /// </summary>
    /// <param name="destination">Caller-owned writable destination, replaced after successful encoding.</param>
    /// <param name="password">Password to open the file; XLS admits 1–15 UTF-16 code units.</param>
    /// <param name="format">Physical workbook format before encryption.</param>
    /// <param name="saveOptions">Format-specific save policy and limits.</param>
    public void SaveEncrypted(Stream destination, string password, ExcelFileFormat format, ExcelSaveOptions? saveOptions = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (password == null) throw new ArgumentNullException(nameof(password));
        if (!destination.CanWrite) throw new ArgumentException("Destination must be writable.", nameof(destination));
        if (format == ExcelFileFormat.Xlsx) { SaveEncrypted(destination, password, saveOptions); return; }
        EnsureWritableForSave();
        byte[] encrypted = BuildEncryptedBinaryWorkbook(format, password, saveOptions);
        WriteEncryptedWorkbookToStream(destination, encrypted);
        LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Encrypted output uses the selected native workbook encoder.");
    }

    /// <summary>Disables plaintext copyback before the first encrypted destination mutation, including failing writes or flushes.</summary>
    private void WriteEncryptedWorkbookToStream(Stream destination, byte[] encrypted) {
        DisableImplicitPersistenceAfterEncryptedSave(destination);
        PrepareDestinationStreamForWrite(destination);
        destination.Write(encrypted, 0, encrypted.Length);
        try { destination.Flush(); } catch (NotSupportedException) { }
        if (destination.CanSeek) destination.Position = 0;
    }

    private bool TrySaveEncryptedBinaryToFile(string path, string password, ExcelSaveOptions? options) {
        ExcelFileFormat? format = ExcelDocumentLoadRouting.HasLegacyXlsExtension(path) ? ExcelFileFormat.Xls
            : ExcelDocumentLoadRouting.HasXlsbExtension(path) ? ExcelFileFormat.Xlsb : null;
        if (format == null) return false;
        EnsureDestinationFileWritable(path); EnsureDirectoryWritable(path);
        byte[] encrypted = BuildEncryptedBinaryWorkbook(format.Value, password, options);
        CommitPreparedPackageToFile(path, encrypted);
        FilePath = path;
        DisableImplicitPersistenceAfterEncryptedSave();
        LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Encrypted output uses the selected native workbook encoder.");
        return true;
    }

    private byte[] BuildEncryptedBinaryWorkbook(ExcelFileFormat format, string password, ExcelSaveOptions? options) {
        long maximumBytes = Math.Min(GetPackageMaterializationLimit(options) ?? int.MaxValue, int.MaxValue);
        if (format == ExcelFileFormat.Xls) {
            LegacyXlsWriter.ValidateEncryptionPassword(password);
            EnsureLegacyXlsSaveDoesNotDropImportedContent(options, includeProjectedChartSheets: true, preserveLinkedVbaProject: true, preserveNativeRecords: true);
            EnsureXlsbStringStorageSupported(isXlsbTarget: false, options);
            if (!CanPreserveNativeLegacyXls(options)) PrepareWorkbookForSave(options);
            byte[] legacy = LegacyXlsWriter.WriteEncryptedWorkbook(this, password, maximumBytes);
            ThrowIfPackageMaterializationExceedsLimit(legacy.LongLength, options);
            return legacy;
        }
        if (format != ExcelFileFormat.Xlsb) throw new ArgumentOutOfRangeException(nameof(format));
        EnsureLegacyXlsSaveDoesNotDropImportedContent(options);
        EnsureXlsbStreamTargetSupported(format, options);
        byte[] package = SourceFormat == ExcelFileFormat.Xlsb
            ? CanCopyUnchangedXlsb(options) ? _xlsbOriginalPackageBytes! : XlsbNativePackageWriter.Rewrite(
                this, _xlsbOriginalPackageBytes!, _xlsbAdvancedWorkbook!, maximumBytes)
            : GenerateBoundedEncryptedXlsbPackage(options, maximumBytes);
        ThrowIfPackageMaterializationExceedsLimit(package.LongLength, options);
        return OfficeEncryption.EncryptPackage(package, password);
    }

    private byte[] GenerateBoundedEncryptedXlsbPackage(ExcelSaveOptions? options, long maximumBytes) {
        PrepareWorkbookForSave(options);
        using var output = new OfficeBoundedMemoryStream(maximumBytes);
        XlsbNewPackageWriter.Write(this, output, options?.XlsbUseSharedStrings == true,
            CancellationToken.None, maximumBytes);
        return output.ToArray();
    }

    /// <summary>Encrypted associated destinations require explicit saves; no password is retained for automatic persistence.</summary>
    private void DisableImplicitPersistenceAfterEncryptedSave(Stream? destination = null) {
        if (destination != null && !ReferenceEquals(destination, _sourceStream)) return;
        _persistenceMode = DocumentPersistenceMode.Explicit;
        _copyPackageToSourceOnDispose = false;
        _copyPackageToFilePathOnDispose = false;
    }
}
