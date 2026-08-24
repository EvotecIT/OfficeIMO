using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

/// <summary>Owns the detection and normalized-read policy applied and reported by the corpus lane.</summary>
internal static class CorpusReaderPolicy {
    internal const int DetectionMaxProbeBytes = 64 * 1024;
    internal const int DetectionMaxContainerEntries = 512;
    internal const int ReadMaxCharacters = 8_000;
    internal const int ReadMaxTableRows = 200;

    internal static ReaderDetectionOptions CreateDetectionOptions() => new() {
        Mode = ReaderDetectionMode.PreferContent,
        MaxProbeBytes = DetectionMaxProbeBytes,
        MaxContainerEntries = DetectionMaxContainerEntries,
        InspectContainers = true
    };

    internal static ReaderOptions CreateReadOptions(long maxInputBytes) => new() {
        MaxInputBytes = maxInputBytes,
        DetectionMode = ReaderDetectionMode.PreferContent,
        DetectionMaxProbeBytes = DetectionMaxProbeBytes,
        DetectionMaxContainerEntries = DetectionMaxContainerEntries,
        ComputeHashes = false,
        MaxChars = ReadMaxCharacters,
        MaxTableRows = ReadMaxTableRows
    };

    internal static CorpusReaderPolicyConfiguration Describe() => new() {
        DetectionMode = ReaderDetectionMode.PreferContent,
        InspectContainers = true,
        DetectionMaxProbeBytes = DetectionMaxProbeBytes,
        DetectionMaxContainerEntries = DetectionMaxContainerEntries,
        ReadMaxCharacters = ReadMaxCharacters,
        ReadMaxTableRows = ReadMaxTableRows,
        ComputeHashes = false
    };
}
