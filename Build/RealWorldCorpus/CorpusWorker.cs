using System.Security.Cryptography;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusWorker {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder()
        .AddAllOfficeIMOHandlers()
        .Build();

    public static CorpusWorkerResult Classify(string path, long maxFileBytes) {
        ValidateFile(path, maxFileBytes);
        string sha256;
        using (FileStream stream = File.OpenRead(path)) {
            sha256 = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        }
        ReaderDetectionResult detection = Reader.Detect(path, new ReaderDetectionOptions {
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

    public static CorpusWorkerResult Probe(string path, long maxFileBytes) {
        ValidateFile(path, maxFileBytes);
        OfficeDocumentReadResult result = Reader.ReadDocument(path, new ReaderOptions {
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

    private static void ValidateFile(string path, long maxFileBytes) {
        var file = new FileInfo(path);
        if (!file.Exists) throw new FileNotFoundException("Corpus input was not found.", path);
        if (file.Length > maxFileBytes) throw new IOException("Corpus input exceeds the configured byte limit.");
    }
}
