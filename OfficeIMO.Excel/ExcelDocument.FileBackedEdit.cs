using System;
using System.IO;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool _usesFileBackedPackage;

        /// <summary>True when the editable Open XML package is staged in a bounded temporary file instead of a byte array.</summary>
        public bool UsesFileBackedPackage => _usesFileBackedPackage;

        /// <summary>
        /// Opens an existing Open XML workbook through a temporary file-backed package. This path avoids buffering the
        /// complete source workbook and checks cancellation while copying. It does not replace the normal optimized load path.
        /// </summary>
        /// <remarks>
        /// The source file is not changed until <see cref="Save()"/> is called, or disposal when
        /// <see cref="DocumentPersistenceMode.SaveOnDispose"/> is selected. XLS and XLSB projection remains available through
        /// <see cref="Load(string, ExcelLoadOptions?)"/>.
        /// </remarks>
        public static ExcelDocument OpenFileBacked(
            string filePath,
            ExcelLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);

            string extension = Path.GetExtension(filePath);
            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase)) {
                throw new NotSupportedException("OpenFileBacked supports Open XML workbooks. Use Load for XLS or XLSB projection.");
            }

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            long? maximumBytes = ResolveInputLimit(resolved);
            long sourceLength = new FileInfo(filePath).Length;
            ValidateFileBackedSourceSize(sourceLength, maximumBytes, resolved);

            cancellationToken.ThrowIfCancellationRequested();
            FileStream? packageStream = null;
            SpreadsheetDocument? package = null;
            try {
                packageStream = OfficeTemporaryFile.Create(
                    "OfficeIMO.Excel-Edit-",
                    extension.Length == 0 ? ".xlsx" : extension,
                    FileOptions.SequentialScan,
                    out _);
                using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete, 81920, FileOptions.SequentialScan)) {
                    CopyFileBackedSource(source, packageStream, maximumBytes, resolved, cancellationToken);
                }

                cancellationToken.ThrowIfCancellationRequested();
                if (resolved.PackageSecurity != null) {
                    packageStream.Position = 0;
                    OfficePackageSecurityInspector.Validate(packageStream, resolved.PackageSecurity);
                    cancellationToken.ThrowIfCancellationRequested();
                }

                packageStream.Position = 0;
                bool normalized = ExcelPackageUtilities.NormalizeContentTypes(packageStream, leaveOpen: true);
                cancellationToken.ThrowIfCancellationRequested();
                packageStream.Position = 0;
                bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
                package = SpreadsheetDocument.Open(packageStream, !readOnly, CreateOpenSettings(resolved.OpenSettings));
                cancellationToken.ThrowIfCancellationRequested();
                ExcelDocument document = CreateDocument(
                    package,
                    filePath,
                    packageStream: null,
                    sourceStream: null,
                    copyPackageToSourceOnDispose: false,
                    leaveSourceStreamOpen: true,
                    copyPackageToFilePathOnDispose: false,
                    ownedOpenStream: packageStream,
                    packageContentTypesKnownNormalized: normalized,
                    unchangedPackageBytes: null,
                    persistenceMode: resolved.PersistenceMode);
                document._usesFileBackedPackage = true;
                cancellationToken.ThrowIfCancellationRequested();
                package = null;
                packageStream = null;
                return document;
            } catch {
                package?.Dispose();
                packageStream?.Dispose();
                throw;
            }
        }

        private static void CopyFileBackedSource(
            Stream source,
            Stream destination,
            long? maximumBytes,
            ExcelLoadOptions options,
            CancellationToken cancellationToken) {
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total = checked(total + read);
                ValidateFileBackedSourceSize(total, maximumBytes, options);
                destination.Write(buffer, 0, read);
            }
            destination.Flush();
            destination.Position = 0;
            cancellationToken.ThrowIfCancellationRequested();
        }

        private static void ValidateFileBackedSourceSize(
            long sourceBytes,
            long? maximumBytes,
            ExcelLoadOptions options) {
            if (!maximumBytes.HasValue || sourceBytes <= maximumBytes.Value) return;
            if (options.PackageSecurity != null
                && maximumBytes.Value == options.PackageSecurity.MaxPackageBytes) {
                OfficePackageSecurityInspector.ValidateSourceSize(sourceBytes, options.PackageSecurity);
            }
            throw new InvalidDataException(
                $"Workbook input contains {sourceBytes} bytes, exceeding MaxInputBytes ({maximumBytes.Value}).");
        }

        private bool TrySaveFileBackedPackageToFile(
            string targetPath,
            ExcelSaveOptions? options,
            CancellationToken cancellationToken) {
            if (!_usesFileBackedPackage) return false;

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            PackagePropertiesSnapshot properties = PackagePropertiesSnapshot.Capture(_spreadSheetDocument);
            long temporaryLimit = ResolveFileBackedTemporaryPackageLimit(options);
            string temporaryPath = string.Empty;
            try {
                using (var stagedFile = CreateTemporarySaveFile(
                    targetPath,
                    FileOptions.SequentialScan,
                    out temporaryPath))
                using (var bounded = new ExcelBoundedSeekableStream(
                    stagedFile,
                    temporaryLimit,
                    leaveOpen: true,
                    cancellationToken))
                using (_spreadSheetDocument.Clone(bounded)) {
                }

                cancellationToken.ThrowIfCancellationRequested();
                properties.ApplyTo(temporaryPath);
                EnsureFileBackedTemporaryPackageWithinLimit(temporaryPath, temporaryLimit);
                cancellationToken.ThrowIfCancellationRequested();
                ExcelPackageUtilities.NormalizeContentTypes(temporaryPath);
                EnsureFileBackedTemporaryPackageWithinLimit(temporaryPath, temporaryLimit);
                ThrowIfOpenXmlValidationFails(temporaryPath, options, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();

                ReplaceTargetFile(temporaryPath, targetPath);
                temporaryPath = string.Empty;
                MarkPackageClean(packageBytes: null);
                FilePath = targetPath;
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("File-backed package finalization avoided managed package materialization.");
                return true;
            } finally {
                DeleteFileIfExists(temporaryPath);
            }
        }

        private static long ResolveFileBackedTemporaryPackageLimit(ExcelSaveOptions? options) {
            long? limit = options == null
                ? ExcelSaveOptions.DefaultMaxTemporaryPackageBytes
                : options.MaxTemporaryPackageBytes;
            if (limit.HasValue && limit.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(ExcelSaveOptions.MaxTemporaryPackageBytes));
            }
            return limit ?? long.MaxValue;
        }

        private static void EnsureFileBackedTemporaryPackageWithinLimit(string path, long limit) {
            long length = new FileInfo(path).Length;
            if (length > limit) {
                throw new InvalidDataException(
                    $"The file-backed Excel package exceeds MaxTemporaryPackageBytes ({limit}).");
            }
        }
    }
}
