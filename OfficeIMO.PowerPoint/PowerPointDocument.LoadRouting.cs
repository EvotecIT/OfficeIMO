using OfficeIMO.Core.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using System.Threading;

namespace OfficeIMO.PowerPoint {
    internal enum LegacyBinaryEncryptionKind {
        NotLegacyBinary,
        Unencrypted,
        Encrypted
    }

    internal static class PowerPointDocumentLoadRouting {
        internal static bool IsLegacyBinary(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                OfficeCompoundDocumentDetector.DocumentKind kind = DetectCompound(bytes, out _);
                switch (kind) {
                    case OfficeCompoundDocumentDetector.DocumentKind.PowerPointDocument:
                        return true;
                    case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                        throw new InvalidDataException("The input contains a legacy Word document, not a PowerPoint presentation. Load it with OfficeIMO.Word.WordDocument.");
                    case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                        throw new InvalidDataException("The input contains a legacy Excel workbook, not a PowerPoint presentation. Load it with OfficeIMO.Excel.ExcelDocument.");
                    case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                        throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use PowerPointDocument.LoadEncrypted and provide its password.");
                    case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                        throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                    default:
                        return HasLegacyBinaryExtension(filePath);
                }
            }
            return HasLegacyBinaryExtension(filePath) && !HasZipSignature(bytes);
        }

        internal static LegacyBinaryEncryptionKind GetLegacyBinaryEncryptionKind(
            byte[] bytes, LegacyPptImportOptions? importOptions = null) {
            if (!OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return LegacyBinaryEncryptionKind.NotLegacyBinary;
            }
            if (DetectCompound(bytes, out _)
                != OfficeCompoundDocumentDetector.DocumentKind
                    .PowerPointDocument) {
                return LegacyBinaryEncryptionKind.NotLegacyBinary;
            }
            LegacyPptImportOptions resolved = importOptions
                ?? new LegacyPptImportOptions();
            bool read;
            OfficeCompoundFile? compound;
            string? error;
            try {
                read = OfficeCompoundFileReader.TryRead(bytes, LegacyPptPackage.CreateCompoundReadOptions(resolved, preserveResourceLimits: true),
                    out compound, out error);
            } catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("powerpoint", exception.LimitName, exception.MaximumValue, exception);
            }
            if (!read
                || compound == null
                || !compound.Streams.TryGetValue("Current User",
                    out byte[]? currentUserStream)) {
                throw new InvalidDataException(error
                    ?? "The binary PowerPoint container has no Current User stream.");
            }

            uint token = LegacyPptCurrentUserAtom.Read(currentUserStream)
                .HeaderToken;
            if (token == LegacyPptCurrentUserAtom.EncryptedHeaderToken) {
                return LegacyBinaryEncryptionKind.Encrypted;
            }
            if (token == LegacyPptCurrentUserAtom.UnencryptedHeaderToken) {
                return LegacyBinaryEncryptionKind.Unencrypted;
            }
            throw new InvalidDataException(
                "The binary PowerPoint Current User stream has an unsupported encryption token.");
        }

        internal static bool HasLegacyBinaryExtension(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".ppt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".pot", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".pps", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".ppa", StringComparison.OrdinalIgnoreCase);
        }

        private static OfficeCompoundDocumentDetector.DocumentKind DetectCompound(byte[] bytes, out string? error) {
            try { return OfficeCompoundDocumentDetector.Detect(bytes, out error, preserveResourceLimits: true); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("powerpoint", exception.LimitName, exception.MaximumValue, exception);
            }
        }

        internal static OfficeCompoundDocumentDetector.DocumentKind DetectCompound(Stream source, long maximumInputBytes,
            int maximumDirectoryEntries, CancellationToken cancellationToken, out string? error) {
            try { return OfficeCompoundDocumentDetector.Detect(source, maximumInputBytes, maximumDirectoryEntries, cancellationToken, out error, preserveResourceLimits: true); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("powerpoint", exception.LimitName, exception.MaximumValue, exception);
            }
        }

        internal static PowerPointFileFormat GetFormat(string? filePath, bool legacyDefault = false) {
            string extension = string.IsNullOrWhiteSpace(filePath) ? string.Empty : Path.GetExtension(filePath);
            if (string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Pptm;
            if (string.Equals(extension, ".pot", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Pot;
            if (string.Equals(extension, ".pps", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Pps;
            if (string.Equals(extension, ".ppa", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Ppt;
            if (string.Equals(extension, ".ppt", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Ppt;
            return legacyDefault ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx;
        }

        private static bool HasZipSignature(byte[] bytes) => bytes.Length >= 2 && bytes[0] == 0x50 && bytes[1] == 0x4B;
    }
}
