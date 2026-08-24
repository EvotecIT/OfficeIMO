using OfficeIMO;

namespace OfficeIMO.RealWorldCorpus;

/// <summary>Owns the package-expansion policy applied and reported by the corpus lane.</summary>
internal static class CorpusPackagePolicy {
    internal const int MaxPartCount = 4_096;
    internal const long MaxPartUncompressedBytes = 64L * 1024L * 1024L;
    internal const long MaxXmlCharactersInPart = 32L * 1024L * 1024L;
    internal const long MaxTotalUncompressedBytes = 256L * 1024L * 1024L;
    internal const double MaxCompressionRatio = 500D;

    internal static OfficePackageSecurityOptions Create(long maxPackageBytes) => new() {
        MaxPackageBytes = maxPackageBytes,
        MaxPartCount = MaxPartCount,
        MaxPartUncompressedBytes = MaxPartUncompressedBytes,
        MaxXmlCharactersInPart = MaxXmlCharactersInPart,
        MaxTotalUncompressedBytes = MaxTotalUncompressedBytes,
        MaxCompressionRatio = MaxCompressionRatio
    };

    internal static CorpusPackagePolicyConfiguration Describe(long maxPackageBytes) => new() {
        MaxPackageBytes = maxPackageBytes,
        MaxPartCount = MaxPartCount,
        MaxPartUncompressedBytes = MaxPartUncompressedBytes,
        MaxXmlCharactersInPart = MaxXmlCharactersInPart,
        MaxTotalUncompressedBytes = MaxTotalUncompressedBytes,
        MaxCompressionRatio = MaxCompressionRatio
    };
}
