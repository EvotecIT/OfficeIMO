using OfficeIMO.Email;
using OfficeIMO.Drawing.Internal;
using System;
using System.IO;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static bool IsOleCompound(byte[] prefix) {
        return StartsWith(prefix, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
    }

    private static DetectionCandidate InspectEmailCompound(byte[] boundedPayload) {
        return EmailDocumentReader.DetectFormat(boundedPayload) == EmailFileFormat.OutlookMsg
            ? DetectionCandidate.High(ReaderInputKind.Email, "application/vnd.ms-outlook",
                "container:msg-properties-stream")
            : DetectionCandidate.Unknown("container:ole-compound-unrecognized");
    }

    private static DetectionCandidate InspectOfficeCompound(byte[] boundedPayload) {
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(boundedPayload, out _));
        return office.Kind != ReaderInputKind.Unknown
            || office.Evidence.Contains(EncryptedOpenXmlEvidence,
                StringComparer.Ordinal)
            ? office
            : InspectEmailCompound(boundedPayload);
    }

    private static DetectionCandidate InspectOfficeCompound(Stream stream,
        long position, int maxContainerEntries) {
        stream.Position = position;
        long remainingBytes = checked(stream.Length - position);
        DetectionCandidate office = MapOfficeCompoundKind(
            OfficeCompoundDocumentDetector.Detect(stream, remainingBytes,
                maxContainerEntries, out _));
        return office.Kind != ReaderInputKind.Unknown
            || office.Evidence.Contains(EncryptedOpenXmlEvidence,
                StringComparer.Ordinal)
            ? office
            : InspectEmailCompound(stream, position, maxContainerEntries);
    }

    private static DetectionCandidate MapOfficeCompoundKind(
        OfficeCompoundDocumentDetector.DocumentKind kind) {
        return kind switch {
            OfficeCompoundDocumentDetector.DocumentKind.WordDocument =>
                DetectionCandidate.High(ReaderInputKind.Word,
                    "application/msword", "container:ole-word-document"),
            OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook =>
                DetectionCandidate.High(ReaderInputKind.Excel,
                    "application/vnd.ms-excel", "container:ole-excel-workbook"),
            OfficeCompoundDocumentDetector.DocumentKind.PowerPointPresentation =>
                DetectionCandidate.High(ReaderInputKind.PowerPoint,
                    "application/vnd.ms-powerpoint",
                    "container:ole-powerpoint-presentation"),
            OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage =>
                DetectionCandidate.Unknown(
                    "container:ole-encrypted-openxml-package"),
            _ => DetectionCandidate.Unknown("container:ole-compound-unrecognized")
        };
    }

    private static DetectionCandidate InspectEmailCompound(Stream stream, long position,
        int maxContainerEntries) {
        stream.Position = position;
        var emailOptions = new EmailReaderOptions(maxCompoundDirectoryEntries: maxContainerEntries);
        return EmailDocumentReader.DetectFormat(stream, emailOptions) == EmailFileFormat.OutlookMsg
            ? DetectionCandidate.High(ReaderInputKind.Email, "application/vnd.ms-outlook",
                "container:msg-properties-stream")
            : DetectionCandidate.Unknown("container:ole-compound-unrecognized");
    }
}
