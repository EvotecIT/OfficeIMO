using System;
using System.Linq;

namespace OfficeIMO.Provenance;

/// <summary>Immutable, provider-specific detector evidence. Scores retain their named scale and are not AI-authorship probabilities.</summary>
public sealed class OfficeProvenanceSignalMeasurement {
    /// <summary>Creates reproducible evidence without retaining the evaluated text or secret watermark keys.</summary>
    /// <param name="detectorVersion">Version of the detector implementation or service.</param>
    /// <param name="algorithm">Name of the watermark algorithm being evaluated.</param>
    /// <param name="scoreName">Provider-defined score scale, such as z-score. Its direction and interpretation belong to the calibration reference.</param>
    /// <param name="score">Finite observed score on that scale.</param>
    /// <param name="threshold">Optional finite decision threshold on the same scale.</param>
    /// <param name="tokenCount">Number of evaluated tokens, when the provider supplies it.</param>
    /// <param name="tokenizer">Tokenizer identity and version, when known.</param>
    /// <param name="configurationId">Non-secret configuration identifier; never include private keys.</param>
    /// <param name="calibrationReference">Versioned calibration reference explaining thresholds, error rates, languages and supported sample lengths. It is not fetched.</param>
    /// <param name="textSha256">SHA-256 of the complete extracted text encoded as UTF-8 before span selection.</param>
    /// <param name="textOffset">Optional zero-based UTF-16 offset in the hashed text; must be paired with length and hash.</param>
    /// <param name="textLength">Optional UTF-16 length of the evaluated span. These are text coordinates, not file mutation offsets.</param>
    public OfficeProvenanceSignalMeasurement(
        string detectorVersion, string algorithm, string scoreName, double score,
        double? threshold = null, int? tokenCount = null, string? tokenizer = null,
        string? configurationId = null, string? calibrationReference = null,
        string? textSha256 = null, int? textOffset = null, int? textLength = null) {
        DetectorVersion = Required(detectorVersion, nameof(detectorVersion));
        Algorithm = Required(algorithm, nameof(algorithm));
        ScoreName = Required(scoreName, nameof(scoreName));
        if (double.IsNaN(score) || double.IsInfinity(score)) throw new ArgumentOutOfRangeException(nameof(score));
        if (threshold.HasValue && (double.IsNaN(threshold.Value) || double.IsInfinity(threshold.Value))) throw new ArgumentOutOfRangeException(nameof(threshold));
        if (tokenCount.HasValue && tokenCount.Value < 0) throw new ArgumentOutOfRangeException(nameof(tokenCount));
        if (textSha256 != null && (textSha256.Length != 64 || !textSha256.All(IsHex))) throw new ArgumentException("A SHA-256 hex digest is required.", nameof(textSha256));
        if (textOffset.HasValue != textLength.HasValue || textOffset.HasValue && textSha256 == null)
            throw new ArgumentException("A text span requires both coordinates and a source text hash.");
        if (textOffset < 0 || textLength < 0 || textOffset.HasValue && textOffset.Value > int.MaxValue - textLength!.Value)
            throw new ArgumentOutOfRangeException(nameof(textOffset), "The text span must fit in UTF-16 string coordinates.");
        Score = score;
        Threshold = threshold;
        TokenCount = tokenCount;
        Tokenizer = tokenizer;
        ConfigurationId = configurationId;
        CalibrationReference = calibrationReference;
        TextSha256 = textSha256?.ToLowerInvariant();
        TextOffset = textOffset;
        TextLength = textLength;
    }

    /// <summary>Gets the implementation or service version.</summary>
    public string DetectorVersion { get; }
    /// <summary>Gets the named watermark algorithm.</summary>
    public string Algorithm { get; }
    /// <summary>Gets the provider-defined score scale.</summary>
    public string ScoreName { get; }
    /// <summary>Gets the finite score, without interpreting it as an authorship probability.</summary>
    public double Score { get; }
    /// <summary>Gets the optional threshold on the provider's score scale.</summary>
    public double? Threshold { get; }
    /// <summary>Gets the evaluated token count.</summary>
    public int? TokenCount { get; }
    /// <summary>Gets the tokenizer identity and version.</summary>
    public string? Tokenizer { get; }
    /// <summary>Gets a non-secret configuration identifier.</summary>
    public string? ConfigurationId { get; }
    /// <summary>Gets the calibration reference; OfficeIMO does not resolve or download it.</summary>
    public string? CalibrationReference { get; }
    /// <summary>Gets the complete extracted-text UTF-8 SHA-256 digest.</summary>
    public string? TextSha256 { get; }
    /// <summary>Gets the UTF-16 offset in the hashed text, not in the encoded file.</summary>
    public int? TextOffset { get; }
    /// <summary>Gets the evaluated UTF-16 span length.</summary>
    public int? TextLength { get; }

    private static string Required(string value, string name) => !string.IsNullOrWhiteSpace(value)
        ? value : throw new ArgumentException("A non-empty value is required.", name);
    private static bool IsHex(char value) => value >= '0' && value <= '9' || value >= 'a' && value <= 'f' || value >= 'A' && value <= 'F';
}
