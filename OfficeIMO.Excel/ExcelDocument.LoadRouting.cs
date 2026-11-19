using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel {
    internal static class ExcelDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };

        internal static bool IsLegacyXls(byte[] bytes, string? filePath) {
            return DetectFormat(bytes, filePath) == ExcelFileFormat.Xls;
        }

        internal static bool IsXlsb(byte[] bytes, string? filePath) {
            return DetectFormat(bytes, filePath) == ExcelFileFormat.Xlsb;
        }

        internal static ExcelFileFormat DetectFormat(byte[] bytes, string? filePath, int? maxPackageEntries = null) {
            ArgumentNullException.ThrowIfNull(bytes);

            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes, encryptedLoad: false, filePath)
                    ? ExcelFileFormat.Xls
                    : ExcelFileFormat.Xlsx;
            }

            if (HasZipSignature(bytes)) {
                int entryLimit = maxPackageEntries ?? OfficePackageSecurityOptions.SecureDefaults.MaxPartCount;
                ValidateZipDirectory(OfficeArchiveSafety.ScanZipCentralDirectory(bytes, entryLimit), entryLimit);

                return XlsbPackageDetector.TryFindWorkbookPart(bytes, entryLimit, out _)
                    ? ExcelFileFormat.Xlsb
                    : ExcelFileFormat.Xlsx;
            }

            if (Legacy.EarlyBiffWorkbook.Identify(bytes) != null || HasLegacyXlsExtension(filePath)) {
                return ExcelFileFormat.Xls;
            }

            return HasXlsbExtension(filePath)
                ? ExcelFileFormat.Xlsb
                : ExcelFileFormat.Xlsx;
        }

        /// <summary>
        /// Applies the ZIP admission that format routing uses to readers that open a ZIP package directly.
        /// Input that does not start with a ZIP signature is left to the reader's own diagnostics.
        /// </summary>
        internal static void ValidateZipPackage(byte[] bytes) {
            if (!HasZipSignature(bytes)) return;
            int entryLimit = OfficePackageSecurityOptions.SecureDefaults.MaxPartCount;
            ValidateZipDirectory(OfficeArchiveSafety.ScanZipCentralDirectory(bytes, entryLimit), entryLimit);
        }

        /// <inheritdoc cref="ValidateZipPackage(byte[])"/>
        internal static void ValidateZipPackage(Stream seekableStream, CancellationToken cancellationToken) {
            long position = seekableStream.Position;
            if (seekableStream.Length - position < ZipSignature.Length) return;
            if (seekableStream.ReadByte() != ZipSignature[0] || seekableStream.ReadByte() != ZipSignature[1]) {
                seekableStream.Position = position;
                return;
            }
            seekableStream.Position = position;
            int entryLimit = OfficePackageSecurityOptions.SecureDefaults.MaxPartCount;
            ValidateZipDirectory(OfficeArchiveSafety.ScanZipCentralDirectory(
                seekableStream, seekableStream.Length - position, entryLimit, cancellationToken), entryLimit);
        }

        private static void ValidateZipDirectory(OfficeArchiveSafety.ZipCentralDirectoryScanResult directory, int entryLimit) {
            // Reject an oversized directory here: either loader would otherwise open the archive
            // and materialize every entry before its own checks run.
            if (directory.LimitExceeded) {
                throw new OfficePackageSecurityException(OfficePackageSecurityRule.PartCount,
                    $"ZIP entry count exceeds the configured maximum of {entryLimit} before package parts are opened.",
                    directory.EntryCount, entryLimit);
            }

            if (directory.HasEncryptedEntries) {
                throw OfficeUnsupportedEncryption.Mark(new InvalidDataException(
                    "The ZIP package contains encrypted entries and cannot be read as a workbook. Password-protected Excel workbooks are encrypted Office packages; open those with ExcelLoadOptions.Password or ExcelReaderOptions.Password."));
            }
        }

        /// <summary>
        /// Distinguishes an encrypted legacy workbook from an encrypted Open XML package by its OLE root streams.
        /// </summary>
        internal static bool IsEncryptedLegacyXls(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes, encryptedLoad: true, filePath);
            }

            return HasLegacyXlsExtension(filePath);
        }

        internal static bool HasLegacyXlsExtension(string? filePath) {
            return !string.IsNullOrWhiteSpace(filePath)
                && string.Equals(Path.GetExtension(filePath), ".xls", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool HasXlsbExtension(string? filePath) {
            return !string.IsNullOrWhiteSpace(filePath)
                && string.Equals(Path.GetExtension(filePath), ".xlsb", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool HasLegacyBinaryExcelExtension(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                return false;
            }

            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xla", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlm", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlw", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasZipSignature(byte[] bytes) {
            return bytes.Length >= ZipSignature.Length
                && bytes[0] == ZipSignature[0]
                && bytes[1] == ZipSignature[1];
        }

        private static bool RouteCompoundDocument(byte[] bytes, bool encryptedLoad, string? filePath) {
            OfficeCompoundDocumentDetector.DocumentKind kind;
            try { kind = OfficeCompoundDocumentDetector.Detect(bytes, out _, preserveResourceLimits: true); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("excel", exception.LimitName, exception.MaximumValue, exception);
            }
            switch (kind) {
                case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                    return true;
                case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                    throw new InvalidDataException("The input contains a legacy Word document, not an Excel workbook. Load it with OfficeIMO.Word.WordDocument.");
                case OfficeCompoundDocumentDetector.DocumentKind.PowerPointDocument:
                    throw new InvalidDataException("The input contains a legacy PowerPoint presentation, not an Excel workbook. Load it with OfficeIMO.PowerPoint.PowerPointDocument.");
                case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                    if (encryptedLoad) {
                        return false;
                    }
                    throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use ExcelDocument.Load with ExcelLoadOptions.Password or ExcelDocument.OpenReader with ExcelReaderOptions.Password.");
                case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                    throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                default:
                    // Normal loads retain their legacy-reader diagnostics. Encrypted loads
                    // fall back to the extension only when the compound root is unknown.
                    return !encryptedLoad || HasLegacyXlsExtension(filePath);
            }
        }
    }
}
