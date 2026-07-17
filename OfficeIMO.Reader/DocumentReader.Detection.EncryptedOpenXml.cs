using OfficeIMO.Drawing.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const string EncryptedOpenXmlEvidence =
        "container:ole-encrypted-openxml-package";

    private static ReaderDetectionResult ResolveEncryptedOpenXmlDetection(
        string path, ReaderOptions options,
        ReaderDetectionResult detection,
        CancellationToken cancellationToken = default) {
        if (!ShouldInspectEncryptedOpenXml(options, detection)) {
            return detection;
        }
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using MemoryStream snapshot = CopyToMemory(stream,
            cancellationToken,
            ResolveInitialMaxInputBytes(path, options));
        return DetectDecryptedOpenXml(snapshot.ToArray(), path, options,
            detection, cancellationToken);
    }

    private static ReaderDetectionResult ResolveEncryptedOpenXmlDetection(
        Stream stream, string? sourceName, ReaderOptions options,
        ReaderDetectionResult detection,
        CancellationToken cancellationToken = default) {
        if (!ShouldInspectEncryptedOpenXml(options, detection)) {
            return detection;
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (!stream.CanSeek) {
            throw new InvalidDataException(
                "Encrypted Open XML content detection requires a seekable stream.");
        }
        long originalPosition = stream.Position;
        try {
            using MemoryStream snapshot = CopyToMemory(stream,
                cancellationToken, options.MaxInputBytes);
            return DetectDecryptedOpenXml(snapshot.ToArray(), sourceName,
                options, detection, cancellationToken);
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static bool ShouldInspectEncryptedOpenXml(
        ReaderOptions options, ReaderDetectionResult detection) =>
        !string.IsNullOrEmpty(options.OpenPassword)
        && detection.Evidence.Contains(EncryptedOpenXmlEvidence,
            StringComparer.Ordinal);

    private static ReaderDetectionResult DetectDecryptedOpenXml(
        byte[] encryptedBytes, string? sourceName, ReaderOptions options,
        ReaderDetectionResult outerDetection,
        CancellationToken cancellationToken) {
        byte[] decrypted = OfficeEncryption.DecryptPackage(encryptedBytes,
            options.OpenPassword!, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        ReaderDetectionResult innerDetection = Detect(decrypted, sourceName,
            CreateDetectionOptions(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (innerDetection.Kind is not (ReaderInputKind.Word
                or ReaderInputKind.Excel
                or ReaderInputKind.PowerPoint)) {
            throw new InvalidDataException(
                "The decrypted Office package does not contain a supported Word, Excel, or PowerPoint document.");
        }

        var evidence = new List<string>(outerDetection.Evidence.Count
            + innerDetection.Evidence.Count + 1);
        evidence.AddRange(outerDetection.Evidence);
        evidence.Add("decryption:open-password");
        evidence.AddRange(innerDetection.Evidence);
        innerDetection.Evidence = evidence.Distinct(
            StringComparer.Ordinal).ToArray();
        innerDetection.ContainerInspected = true;
        innerDetection.ContentInspected = true;
        innerDetection.InspectedBytes = Math.Max(
            outerDetection.InspectedBytes,
            innerDetection.InspectedBytes);
        return innerDetection;
    }
}
