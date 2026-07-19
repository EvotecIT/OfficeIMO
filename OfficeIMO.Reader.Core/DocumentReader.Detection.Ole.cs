using OfficeIMO.Reader.Internal.Compound;
using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const string EncryptedOpenXmlEvidence = "container:ole-encrypted-openxml-package";

    private static bool IsOleCompound(byte[] prefix) => StartsWith(
        prefix,
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });

    private static DetectionCandidate InspectOfficeCompound(byte[] boundedPayload) {
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(boundedPayload, out _));
        if (office.Kind != ReaderInputKind.Unknown ||
            office.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal)) {
            return office;
        }

        using var stream = new MemoryStream(boundedPayload, writable: false);
        return InspectMsgCompound(stream, 0, boundedPayload.LongLength,
            ReaderOptions.DefaultDetectionMaxContainerEntries, CancellationToken.None);
    }

    private static DetectionCandidate InspectOfficeCompound(
        Stream stream,
        long position,
        int maxContainerEntries,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        stream.Position = position;
        long remainingBytes = checked(stream.Length - position);
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(stream, remainingBytes,
                maxContainerEntries, cancellationToken, out _));
        if (office.Kind != ReaderInputKind.Unknown ||
            office.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal)) {
            return office;
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
            OfficeCompoundDocumentDetector.DocumentKind.PowerPointPresentation =>
                DetectionCandidate.High(ReaderInputKind.PowerPoint, "application/vnd.ms-powerpoint",
                    "container:ole-powerpoint-presentation"),
            OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage =>
                DetectionCandidate.Unknown(EncryptedOpenXmlEvidence),
            _ => DetectionCandidate.Unknown("container:ole-compound-unrecognized")
        };

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
                maxContainerEntries, cancellationToken, out bool contains, out _);
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
