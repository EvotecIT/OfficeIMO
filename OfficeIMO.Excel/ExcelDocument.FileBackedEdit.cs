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
            if (maximumBytes.HasValue && sourceLength > maximumBytes.Value) {
                throw new InvalidDataException(
                    $"Workbook input contains {sourceLength} bytes, exceeding MaxInputBytes ({maximumBytes.Value}).");
            }

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
                    CopyFileBackedSource(source, packageStream, maximumBytes, cancellationToken);
                }

                if (resolved.PackageSecurity != null) {
                    packageStream.Position = 0;
                    OfficePackageSecurityInspector.Validate(packageStream, resolved.PackageSecurity);
                }

                packageStream.Position = 0;
                bool normalized = ExcelPackageUtilities.NormalizeContentTypes(packageStream, leaveOpen: true);
                packageStream.Position = 0;
                bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
                package = SpreadsheetDocument.Open(packageStream, !readOnly, CreateOpenSettings(resolved.OpenSettings));
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
            CancellationToken cancellationToken) {
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total = checked(total + read);
                if (maximumBytes.HasValue && total > maximumBytes.Value) {
                    throw new InvalidDataException(
                        $"Workbook input exceeds MaxInputBytes ({maximumBytes.Value}).");
                }
                destination.Write(buffer, 0, read);
            }
            destination.Flush();
            destination.Position = 0;
            cancellationToken.ThrowIfCancellationRequested();
        }
    }
}
