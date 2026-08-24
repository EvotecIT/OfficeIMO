using System.Security.Cryptography;
using OfficeIMO;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusWorker {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder()
        .AddAllOfficeIMOHandlers()
        .Build();

    public static CorpusWorkerResult Classify(string path, long maxFileBytes) {
        byte[] snapshot = ReadBoundedSnapshot(path, maxFileBytes);
        string sha256 = Convert.ToHexString(SHA256.HashData(snapshot)).ToLowerInvariant();
        ReaderDetectionResult detection = Reader.Detect(snapshot, Path.GetFileName(path), new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent,
            MaxProbeBytes = 64 * 1024,
            MaxContainerEntries = 512,
            InspectContainers = true
        });
        return new CorpusWorkerResult {
            Stage = CorpusOutcomes.Classification,
            Succeeded = true,
            Sha256 = sha256,
            ExtensionKind = detection.ExtensionKind,
            ContentKind = detection.ContentKind,
            ContentConfidence = detection.ContentConfidence,
            DetectedKind = detection.Kind,
            Confidence = detection.Confidence,
            IsMismatch = detection.IsMismatch,
            Evidence = detection.Evidence
        };
    }

    public static CorpusWorkerResult Probe(string path, long maxFileBytes, string expectedSha256) {
        byte[] snapshot = ReadBoundedSnapshot(path, maxFileBytes);
        string actualSha256 = Convert.ToHexString(SHA256.HashData(snapshot)).ToLowerInvariant();
        if (!string.Equals(actualSha256, expectedSha256, StringComparison.Ordinal)) {
            throw new CorpusInputChangedException();
        }
        ValidatePackage(snapshot, Path.GetFileName(path), maxFileBytes);
        OfficeDocumentReadResult result = Reader.ReadDocument(snapshot, Path.GetFileName(path), new ReaderOptions {
            MaxInputBytes = maxFileBytes,
            DetectionMode = ReaderDetectionMode.PreferContent,
            DetectionMaxProbeBytes = 64 * 1024,
            DetectionMaxContainerEntries = 512,
            ComputeHashes = false,
            MaxChars = 8_000,
            MaxTableRows = 200
        });
        return new CorpusWorkerResult {
            Stage = CorpusOutcomes.Probe,
            Succeeded = true,
            DetectedKind = result.Kind,
            ChunkCount = result.Chunks.Count,
            PageCount = result.Pages.Count,
            BlockCount = result.Blocks.Count,
            AssetCount = result.Assets.Count,
            InformationDiagnostics = result.Diagnostics.Count(item => item.Severity == OfficeDocumentDiagnosticSeverity.Information),
            WarningDiagnostics = result.Diagnostics.Count(item => item.Severity == OfficeDocumentDiagnosticSeverity.Warning),
            ErrorDiagnostics = result.Diagnostics.Count(item => item.Severity == OfficeDocumentDiagnosticSeverity.Error),
            DiagnosticCodes = result.Diagnostics.Select(item => item.Code)
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(code => code, StringComparer.Ordinal)
                .ToArray()
        };
    }

    private static void ValidatePackage(byte[] snapshot, string sourceName, long maxFileBytes) {
        ReaderDetectionResult detection = Reader.Detect(snapshot, sourceName, new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent,
            MaxProbeBytes = 64 * 1024,
            MaxContainerEntries = 512,
            InspectContainers = true
        });
        if (!IsPackageKind(detection.ContentKind) &&
            !IsPackageKind(detection.Kind)) {
            return;
        }

        OfficePackageSecurityInspector.Validate(snapshot, CorpusPackagePolicy.Create(maxFileBytes));
    }

    private static bool IsPackageKind(ReaderInputKind kind) =>
        kind is ReaderInputKind.Word or ReaderInputKind.Excel or ReaderInputKind.PowerPoint;

    private static byte[] ReadBoundedSnapshot(string path, long maxFileBytes) {
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var snapshot = new MemoryStream((int)Math.Min(maxFileBytes, 1024L * 1024L));
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maxFileBytes) {
                throw new IOException("Corpus input exceeds the configured byte limit.");
            }
            snapshot.Write(buffer, 0, read);
        }
        return snapshot.ToArray();
    }
}

internal sealed class CorpusInputChangedException : IOException {
    internal CorpusInputChangedException()
        : base("Corpus input changed after classification; the recorded hash was not parsed.") { }
}
