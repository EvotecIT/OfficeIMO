using OfficeIMO.Reader.Internal.Compound;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const string EncryptedOpenXmlEvidence = "container:ole-encrypted-openxml-package";

    private static bool IsOleCompound(byte[] prefix) => StartsWith(
        prefix,
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });

    private static DetectionCandidate InspectOfficeCompound(byte[] boundedPayload) {
        try {
            return InspectOfficeCompoundBytes(boundedPayload);
        } catch (OfficeCompoundStreamLimitExceededException exception) {
            throw CompoundDetectionLimit(exception);
        }
    }

    private static DetectionCandidate InspectOfficeCompoundBytes(byte[] boundedPayload) {
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(boundedPayload, out _, preserveResourceLimits: true));
        using var stream = new MemoryStream(boundedPayload, writable: false);
        if (office.Kind != ReaderInputKind.Unknown ||
            office.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal)) {
            return ApplyLegacyBinaryEncryption(office, stream, 0,
                ReaderOptions.DefaultDetectionMaxContainerEntries, CancellationToken.None);
        }

        return InspectMsgCompound(stream, 0, boundedPayload.LongLength,
            ReaderOptions.DefaultDetectionMaxContainerEntries, CancellationToken.None);
    }

    private static DetectionCandidate InspectOfficeCompound(
        Stream stream,
        long position,
        int maxContainerEntries,
        CancellationToken cancellationToken = default) {
        try {
            return InspectOfficeCompoundStream(stream, position, maxContainerEntries, cancellationToken);
        } catch (OfficeCompoundStreamLimitExceededException exception) {
            throw CompoundDetectionLimit(exception);
        }
    }

    private static ReaderResourceLimitException CompoundDetectionLimit(
        OfficeCompoundStreamLimitExceededException exception) =>
        new ReaderResourceLimitException("Detection." + exception.LimitName,
            exception.MaximumValue, exception);

    private static DetectionCandidate InspectOfficeCompoundStream(
        Stream stream,
        long position,
        int maxContainerEntries,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        stream.Position = position;
        long remainingBytes = checked(stream.Length - position);
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(stream, remainingBytes,
                maxContainerEntries, cancellationToken, out _, preserveResourceLimits: true));
        if (office.Kind != ReaderInputKind.Unknown ||
            office.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal)) {
            return ApplyLegacyBinaryEncryption(office, stream, position, maxContainerEntries, cancellationToken);
        }

        return InspectMsgCompound(stream, position, remainingBytes,
            maxContainerEntries, cancellationToken);
    }

    private static DetectionCandidate MapOfficeCompoundKind(
        OfficeCompoundDocumentDetector.DocumentKind kind) => kind switch {
            OfficeCompoundDocumentDetector.DocumentKind.WordDocument =>
                DetectionCandidate.High(ReaderInputKind.Word, "application/msword",
                    "container:ole-word-document"),
            OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook =>
                DetectionCandidate.High(ReaderInputKind.Excel, "application/vnd.ms-excel",
                    "container:ole-excel-workbook"),
            OfficeCompoundDocumentDetector.DocumentKind.PowerPointDocument =>
                DetectionCandidate.High(ReaderInputKind.PowerPoint, "application/vnd.ms-powerpoint",
                    "container:ole-powerpoint-presentation"),
            OfficeCompoundDocumentDetector.DocumentKind.VisioDrawing =>
                DetectionCandidate.High(ReaderInputKind.Visio, "application/vnd.visio", "container:ole-visio-document"),
            OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage =>
                DetectionCandidate.Unknown(EncryptedOpenXmlEvidence).WithEncryption("encryption:office-encrypted-package"),
            OfficeCompoundDocumentDetector.DocumentKind.PublisherPublication =>
                DetectionCandidate.High(ReaderInputKind.Publisher, "application/x-mspublisher",
                    "container:ole-publisher-quill"),
            _ => DetectionCandidate.Unknown("container:ole-compound-unrecognized")
        };

    private const ushort WordFibEncryptedFlag = 0x0100;
    private const ushort ExcelBofRecord = 0x0809;
    private const ushort ExcelFilePassRecord = 0x002F;
    private const uint PowerPointEncryptedCurrentUserToken = 0xF3D1C4DF;
    private const int LegacyEncryptionProbeBytes = 4096;

    /// <summary>
    /// Marks a legacy Word, Excel or PowerPoint file encrypted when its own header says so: the Word FIB fEncrypted bit,
    /// an Excel FILEPASS record after the workbook BOF, or the encrypted PowerPoint Current User token. Only a bounded
    /// prefix of one compound stream is read; a stream that cannot be inspected leaves the candidate unchanged.
    /// </summary>
    private static DetectionCandidate ApplyLegacyBinaryEncryption(DetectionCandidate office, Stream stream, long position,
        int maxContainerEntries, CancellationToken cancellationToken) {
        if (office.IsEncrypted) return office;
        string[] names;
        string evidence;
        switch (office.Kind) {
            case ReaderInputKind.Word:
                names = new[] { "WordDocument" };
                evidence = "encryption:legacy-word-fib";
                break;
            case ReaderInputKind.Excel:
                names = new[] { "Workbook", "Book" };
                evidence = "encryption:legacy-excel-filepass";
                break;
            case ReaderInputKind.PowerPoint:
                names = new[] { "Current User" };
                evidence = "encryption:legacy-powerpoint-current-user";
                break;
            default:
                return office;
        }

        long originalPosition = stream.Position;
        try {
            stream.Position = position;
            var options = new OfficeCompoundReadOptions(maxDirectoryEntries: Math.Max(1, maxContainerEntries));
            if (!OfficeCompoundFileReader.TryOpenStream(stream, options,
                    (path, _) => names.Contains(path, StringComparer.OrdinalIgnoreCase), leaveOpen: true,
                    cancellationToken, out Stream? selected, out _) || selected == null) {
                return office;
            }

            byte[] prefix;
            using (selected) {
                prefix = ReadDetectionPrefix(selected, LegacyEncryptionProbeBytes);
            }
            bool encrypted = office.Kind switch {
                ReaderInputKind.Word => prefix.Length >= 12 && (ReadUInt16(prefix, 10) & WordFibEncryptedFlag) != 0,
                ReaderInputKind.Excel => HasExcelFilePassRecord(prefix),
                _ => prefix.Length >= 16 && ReadUInt32(prefix, 12) == PowerPointEncryptedCurrentUserToken
            };
            return encrypted ? office.WithEncryption(evidence) : office;
        } catch (InvalidDataException) {
            return office;
        } catch (EndOfStreamException) {
            return office;
        } catch (OfficeCompoundStreamLimitExceededException) {
            return office;
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static bool HasExcelFilePassRecord(byte[] workbook) {
        // FILEPASS belongs to the workbook globals and follows its BOF ahead of the sheet records.
        if (workbook.Length < 4 || ReadUInt16(workbook, 0) != ExcelBofRecord) return false;
        int offset = 0;
        for (int record = 0; record < 16 && offset + 4 <= workbook.Length; record++) {
            ushort type = ReadUInt16(workbook, offset);
            if (type == ExcelFilePassRecord) return true;
            offset += 4 + ReadUInt16(workbook, offset + 2);
        }
        return false;
    }

    private static DetectionCandidate InspectMsgCompound(
        Stream stream,
        long position,
        long maxInputBytes,
        int maxContainerEntries,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long originalPosition = stream.Position;
        try {
            stream.Position = position;
            bool inspected = OfficeCompoundFileReader.TryContainsStreamPath(
                stream, "__properties_version1.0", Math.Max(512, maxInputBytes),
                maxContainerEntries, cancellationToken, out bool contains, out _, preserveResourceLimits: true);
            return inspected && contains
                ? DetectionCandidate.High(ReaderInputKind.Email,
                    "application/vnd.ms-outlook", "container:msg-properties-stream")
                : DetectionCandidate.Unknown("container:ole-compound-unrecognized");
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static ReaderDetectionResult ResolveEncryptedOpenXmlDetection(
        string path,
        ReaderOptions options,
        ReaderDetectionResult detection,
        CancellationToken cancellationToken = default) {
        if (!ShouldProbeEncryptedOpenXml(options, detection)) return detection;
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        return ProbeEncryptedOpenXml(stream, path, options, detection, cancellationToken);
    }

    private static ReaderDetectionResult ResolveEncryptedOpenXmlDetection(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        ReaderDetectionResult detection,
        CancellationToken cancellationToken = default) {
        if (!ShouldProbeEncryptedOpenXml(options, detection)) return detection;
        cancellationToken.ThrowIfCancellationRequested();
        if (!stream.CanSeek) {
            throw new InvalidDataException("Encrypted Open XML content probing requires a seekable stream.");
        }
        return ProbeEncryptedOpenXml(stream, sourceName, options, detection, cancellationToken);
    }

    private static bool ShouldProbeEncryptedOpenXml(ReaderOptions options, ReaderDetectionResult detection) =>
        !string.IsNullOrEmpty(options.OpenPassword) &&
        detection.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal);

    private static ReaderDetectionResult ProbeEncryptedOpenXml(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        ReaderDetectionResult detection,
        CancellationToken cancellationToken) {
        long position = stream.Position;
        try {
            if (!GetActiveHandlerRegistry().TryProbeStream(stream, sourceName, options,
                    cancellationToken, out ReaderHandlerDescriptor handler)) {
                return detection;
            }
            detection.ContentKind = handler.Kind;
            detection.Kind = handler.Kind;
            detection.ContentConfidence = ReaderDetectionConfidence.High;
            detection.Confidence = ReaderDetectionConfidence.High;
            detection.ContentInspected = true;
            detection.ContainerInspected = true;
            detection.Evidence = detection.Evidence
                .Concat(new[] { "decryption:open-password", "handler-probe:" + handler.Id })
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            return detection;
        } finally {
            stream.Position = position;
        }
    }
}
