using OfficeIMO.Email;
using System.IO;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static bool IsOleCompound(byte[] prefix) {
        return StartsWith(prefix, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
    }

    private static DetectionCandidate InspectEmailCompound(byte[] boundedPayload) {
        return EmailDocumentReader.DetectFormat(boundedPayload) == EmailFileFormat.OutlookMsg
            ? DetectionCandidate.High(ReaderInputKind.Email, "application/vnd.ms-outlook",
                "container:msg-properties-stream")
            : DetectionCandidate.Unknown("container:ole-compound-unrecognized");
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
