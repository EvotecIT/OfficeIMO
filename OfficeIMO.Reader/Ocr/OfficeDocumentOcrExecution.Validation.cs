using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private static List<CandidateJob> BuildJobs(
        OfficeDocumentReadResult document,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates,
        IReadOnlyList<OfficeDocumentAsset> assets,
        OfficeOcrEngineCapabilities capabilities,
        string engineId,
        ExecutionOptionsSnapshot options,
        List<OfficeDocumentDiagnostic> diagnostics) {
        var jobs = new List<CandidateJob>(Math.Min(candidates.Count, options.MaxCandidates));
        if (candidates.Count > options.MaxCandidates) {
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "ocr-candidate-limit",
                Message = "OCR candidates were limited to MaxCandidates (" + options.MaxCandidates + ").",
                Source = engineId,
                IsRecoverable = true,
                Location = document.Source == null ? null : new ReaderLocation { Path = document.Source.Path },
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["candidateCount"] = candidates.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["selectedCount"] = options.MaxCandidates.ToString(System.Globalization.CultureInfo.InvariantCulture)
                }
            });
        }

        long totalBytes = 0;
        int selectedCount = Math.Min(candidates.Count, options.MaxCandidates);
        for (int index = 0; index < selectedCount; index++) {
            OfficeDocumentOcrCandidate candidate = candidates[index];
            OfficeDocumentAsset? asset = ResolveAsset(candidate, assets, out string? resolutionCode);
            if (asset == null) {
                diagnostics.Add(BuildDiagnostic(
                    candidate,
                    null,
                    engineId,
                    OfficeDocumentDiagnosticSeverity.Warning,
                    OfficeDocumentDiagnosticCategory.Ocr,
                    resolutionCode ?? "ocr-asset-missing",
                    resolutionCode == "ocr-asset-ambiguous"
                        ? "The OCR candidate does not identify one unambiguous source asset."
                        : "The OCR candidate's source asset was not found.",
                    true));
                continue;
            }

            byte[]? sourcePayload = asset.PayloadBytes;
            if (sourcePayload == null || sourcePayload.Length == 0) {
                diagnostics.Add(BuildDiagnostic(candidate, asset, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Ocr,
                    "ocr-asset-payload-missing", "The OCR source asset has no materialized payload bytes.", true));
                continue;
            }
            if (sourcePayload.LongLength > options.MaxInputBytesPerCandidate) {
                diagnostics.Add(BuildDiagnostic(candidate, asset, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                    "ocr-input-limit", "The OCR source asset exceeds MaxInputBytesPerCandidate.", true,
                    BuildLimitAttributes(sourcePayload.LongLength, options.MaxInputBytesPerCandidate)));
                continue;
            }
            if (sourcePayload.LongLength > options.MaxTotalInputBytes - totalBytes) {
                diagnostics.Add(BuildDiagnostic(candidate, asset, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                    "ocr-total-input-limit", "The OCR source asset was skipped because MaxTotalInputBytes was reached.", true,
                    BuildLimitAttributes(totalBytes + sourcePayload.LongLength, options.MaxTotalInputBytes)));
                continue;
            }
            if (!IsSupportedMediaType(asset.MediaType, capabilities.SupportedMediaTypes)) {
                string mediaType = string.IsNullOrWhiteSpace(asset.MediaType) ? "(unknown)" : asset.MediaType!;
                diagnostics.Add(BuildDiagnostic(candidate, asset, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Ocr,
                    "ocr-media-type-unsupported", "The OCR engine does not advertise support for media type '" + mediaType + "'.", true));
                continue;
            }
            if (options.RequirePayloadHashMatch && !string.IsNullOrWhiteSpace(asset.PayloadHash) && !asset.PayloadHashMatches(out string? actualHash)) {
                diagnostics.Add(BuildDiagnostic(candidate, asset, engineId, OfficeDocumentDiagnosticSeverity.Error, OfficeDocumentDiagnosticCategory.Input,
                    "ocr-payload-hash-mismatch", "The OCR source asset payload does not match its declared hash.", false,
                    new Dictionary<string, string>(StringComparer.Ordinal) { ["actualHash"] = actualHash ?? string.Empty }));
                continue;
            }

            byte[] payload = sourcePayload.ToArray();
            jobs.Add(new CandidateJob(index, candidate, asset, payload));
            totalBytes += payload.LongLength;
        }
        return jobs;
    }

    private static OfficeDocumentAsset? ResolveAsset(
        OfficeDocumentOcrCandidate candidate,
        IReadOnlyList<OfficeDocumentAsset> assets,
        out string? resolutionCode) {
        resolutionCode = null;
        if (!string.IsNullOrWhiteSpace(candidate.AssetId)) {
            OfficeDocumentAsset? exact = assets.FirstOrDefault(asset => string.Equals(asset.Id, candidate.AssetId, StringComparison.Ordinal));
            if (exact == null) resolutionCode = "ocr-asset-missing";
            if (exact != null && IsAmbiguousMultiImagePage(candidate)) {
                resolutionCode = "ocr-asset-ambiguous";
                return null;
            }
            return exact;
        }

        OfficeDocumentAsset[] matches = assets
            .Where(static asset => string.Equals(asset.Kind, "image", StringComparison.OrdinalIgnoreCase))
            .Where(asset => IsSameContainer(candidate.Location, asset.Location))
            .Take(2)
            .ToArray();
        if (matches.Length == 1) return matches[0];
        resolutionCode = matches.Length == 0 ? "ocr-asset-missing" : "ocr-asset-ambiguous";
        return null;
    }

    private static bool IsAmbiguousMultiImagePage(OfficeDocumentOcrCandidate candidate) {
        return candidate.ImageCount.GetValueOrDefault() > 1
            && string.Equals(candidate.Kind, "page", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSameContainer(ReaderLocation candidate, ReaderLocation asset) {
        if (candidate.Page.HasValue) return candidate.Page == asset.Page;
        if (candidate.Slide.HasValue) return candidate.Slide == asset.Slide;
        if (!string.IsNullOrWhiteSpace(candidate.Sheet)) return string.Equals(candidate.Sheet, asset.Sheet, StringComparison.Ordinal);
        if (!string.IsNullOrWhiteSpace(candidate.A1Range)) return string.Equals(candidate.A1Range, asset.A1Range, StringComparison.Ordinal);
        return string.Equals(candidate.Path, asset.Path, StringComparison.Ordinal);
    }

    private static bool IsSupportedMediaType(string? mediaType, IReadOnlyList<string>? supported) {
        if (supported == null || supported.Count == 0) return true;
        if (string.IsNullOrWhiteSpace(mediaType)) return false;
        foreach (string declared in supported) {
            if (string.IsNullOrWhiteSpace(declared)) continue;
            if (string.Equals(declared, mediaType, StringComparison.OrdinalIgnoreCase)) return true;
            if (declared.EndsWith("/*", StringComparison.Ordinal) && mediaType!.StartsWith(declared.Substring(0, declared.Length - 1), StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static void NormalizeEngineResult(
        OfficeOcrEngineResult result,
        string engineId,
        ExecutionOptionsSnapshot options,
        OfficeDocumentOcrCandidate candidate,
        List<OfficeDocumentDiagnostic> executionDiagnostics) {
        result.Text = result.Text?.Trim() ?? string.Empty;
        if (result.Text.Length > options.MaxRecognizedCharactersPerCandidate) {
            result.Text = TruncateText(result.Text, options.MaxRecognizedCharactersPerCandidate);
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-text-limit", "OCR recognized text was truncated at MaxRecognizedCharactersPerCandidate.", true));
        }
        result.Provider = string.IsNullOrWhiteSpace(result.Provider) ? engineId : result.Provider!.Trim();
        result.Language = string.IsNullOrWhiteSpace(result.Language) ? options.Language : result.Language!.Trim();
        bool adjustedConfidence = false;
        result.Confidence = NormalizeConfidence(result.Confidence, ref adjustedConfidence);
        OfficeOcrTextSpan[] sourceSpans = (result.Spans ?? Array.Empty<OfficeOcrTextSpan>())
            .Where(static span => span != null)
            .OrderBy(static span => span.Sequence)
            .ThenBy(static span => span.Level)
            .ToArray();
        if (sourceSpans.Length > options.MaxSpansPerCandidate) {
            sourceSpans = sourceSpans.Take(options.MaxSpansPerCandidate).ToArray();
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-span-limit", "OCR detailed spans were truncated at MaxSpansPerCandidate.", true));
        }
        OfficeOcrTextSpan[] spans = sourceSpans;
        for (int index = 0; index < spans.Length; index++) {
            OfficeOcrTextSpan span = spans[index];
            span.Sequence = index;
            span.Text = span.Text?.Trim() ?? string.Empty;
            span.Language = string.IsNullOrWhiteSpace(span.Language) ? result.Language : span.Language!.Trim();
            span.Confidence = NormalizeConfidence(span.Confidence, ref adjustedConfidence);
        }
        result.Spans = spans;
        if (adjustedConfidence) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Ocr,
                "ocr-confidence-out-of-range", "One or more OCR confidence values were normalized; non-finite values were removed and out-of-range values were clamped.", true));
        }

        OfficeDocumentDiagnostic[] providerDiagnostics = (result.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>())
            .Where(static diagnostic => diagnostic != null)
            .ToArray();
        foreach (OfficeDocumentDiagnostic diagnostic in providerDiagnostics) {
            if (diagnostic.Category == OfficeDocumentDiagnosticCategory.General) diagnostic.Category = OfficeDocumentDiagnosticCategory.Ocr;
            if (string.IsNullOrWhiteSpace(diagnostic.Source)) diagnostic.Source = engineId;
            if (diagnostic.Location == null) diagnostic.Location = candidate.Location;
        }
        result.Diagnostics = providerDiagnostics;
    }

    private static double? NormalizeConfidence(double? value, ref bool adjusted) {
        if (!value.HasValue) return null;
        if (double.IsNaN(value.Value) || double.IsInfinity(value.Value)) {
            adjusted = true;
            return null;
        }
        if (value.Value >= 0D && value.Value <= 1D) return value;
        adjusted = true;
        return value.Value < 0D ? 0D : 1D;
    }

    private static string TruncateText(string value, int maxCharacters) {
        int length = maxCharacters;
        if (length > 0 && length < value.Length && char.IsHighSurrogate(value[length - 1]) && char.IsLowSurrogate(value[length])) length--;
        return value.Substring(0, length);
    }

    private static Dictionary<string, string> BuildLimitAttributes(long actual, long limit) {
        return new Dictionary<string, string>(StringComparer.Ordinal) {
            ["actualBytes"] = actual.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["limitBytes"] = limit.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };
    }

    private static OfficeDocumentDiagnostic BuildDiagnostic(
        OfficeDocumentOcrCandidate candidate,
        OfficeDocumentAsset? asset,
        string source,
        OfficeDocumentDiagnosticSeverity severity,
        OfficeDocumentDiagnosticCategory category,
        string code,
        string message,
        bool recoverable,
        IReadOnlyDictionary<string, string>? attributes = null) {
        var details = attributes == null
            ? new Dictionary<string, string>(StringComparer.Ordinal)
            : attributes.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal);
        details["candidateId"] = candidate.Id;
        if (asset != null) details["assetId"] = asset.Id;
        return new OfficeDocumentDiagnostic {
            Severity = severity,
            Category = category,
            Code = code,
            Message = message,
            Source = source,
            IsRecoverable = recoverable,
            Location = candidate.Location,
            Attributes = details
        };
    }
}
