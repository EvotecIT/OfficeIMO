using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Ocr;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private static List<CandidateJob> BuildJobs(
        OfficeDocumentReadResult document,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates,
        IReadOnlyList<OfficeDocumentAsset> assets,
        OcrEngineCapabilities capabilities,
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
        OcrResult result,
        string engineId,
        ExecutionOptionsSnapshot options,
        OfficeDocumentOcrCandidate candidate,
        List<OfficeDocumentDiagnostic> executionDiagnostics) {
        bool truncatedRecognizedText = false;
        int remainingRecognizedCharacters = options.MaxRecognizedCharactersPerCandidate;
        result.Text = ConsumeRequiredText(result.Text, ref remainingRecognizedCharacters, ref truncatedRecognizedText);
        if (truncatedRecognizedText) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-text-limit", "OCR recognized text was truncated at MaxRecognizedCharactersPerCandidate.", true));
        }
        bool truncatedResultMetadata = false;
        int remainingResultMetadataCharacters = options.MaxResultMetadataCharactersPerCandidate;
        result.Provider = ConsumeOptionalText(result.Provider, ref remainingResultMetadataCharacters, ref truncatedResultMetadata);
        if (result.Provider == null) {
            result.Provider = ConsumeOptionalText(engineId, ref remainingResultMetadataCharacters, ref truncatedResultMetadata);
        }
        result.Model = ConsumeOptionalText(result.Model, ref remainingResultMetadataCharacters, ref truncatedResultMetadata);
        result.Language = ConsumeOptionalText(result.Language, ref remainingResultMetadataCharacters, ref truncatedResultMetadata);
        if (result.Language == null) {
            result.Language = ConsumeOptionalText(options.Language, ref remainingResultMetadataCharacters, ref truncatedResultMetadata);
        }
        if (truncatedResultMetadata) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-result-metadata-limit", "OCR result metadata was truncated at MaxResultMetadataCharactersPerCandidate.", true));
        }
        bool adjustedConfidence = false;
        bool discardedHierarchyId = false;
        result.Confidence = NormalizeConfidence(result.Confidence, ref adjustedConfidence);
        IReadOnlyList<OcrTextSpan> returnedSpans = result.Spans ?? Array.Empty<OcrTextSpan>();
        int spanLimit = Math.Min(returnedSpans.Count, options.MaxSpansPerCandidate);
        var boundedSpans = new List<OcrTextSpan>(spanLimit);
        for (int index = 0; index < spanLimit; index++) {
            OcrTextSpan? span = returnedSpans[index];
            if (span != null) boundedSpans.Add(span);
        }
        OcrTextSpan[] spans = boundedSpans
            .OrderBy(static span => span.Sequence)
            .ThenBy(static span => span.Level)
            .ToArray();
        if (returnedSpans.Count > options.MaxSpansPerCandidate) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-span-limit", "OCR detailed spans were truncated at MaxSpansPerCandidate.", true));
        }
        bool truncatedSpanCharacters = false;
        int remainingSpanCharacters = options.MaxSpanCharactersPerCandidate;
        for (int index = 0; index < spans.Length; index++) {
            OcrTextSpan span = spans[index];
            span.Sequence = index;
            span.Text = ConsumeRequiredText(span.Text, ref remainingSpanCharacters, ref truncatedSpanCharacters);
            span.Language = ConsumeOptionalText(span.Language, ref remainingSpanCharacters, ref truncatedSpanCharacters);
            if (span.Language == null) {
                span.Language = ConsumeOptionalText(result.Language, ref remainingSpanCharacters, ref truncatedSpanCharacters);
            }
            span.BlockId = ConsumeOptionalText(
                NormalizeHierarchyId(span.BlockId, ref discardedHierarchyId),
                ref remainingSpanCharacters,
                ref truncatedSpanCharacters);
            span.ParagraphId = ConsumeOptionalText(
                NormalizeHierarchyId(span.ParagraphId, ref discardedHierarchyId),
                ref remainingSpanCharacters,
                ref truncatedSpanCharacters);
            span.LineId = ConsumeOptionalText(
                NormalizeHierarchyId(span.LineId, ref discardedHierarchyId),
                ref remainingSpanCharacters,
                ref truncatedSpanCharacters);
            span.Confidence = NormalizeConfidence(span.Confidence, ref adjustedConfidence);
        }
        result.Spans = spans;
        if (truncatedSpanCharacters) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-span-text-limit", "OCR span text and metadata were truncated at MaxSpanCharactersPerCandidate.", true));
        }
        if (adjustedConfidence) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Ocr,
                "ocr-confidence-out-of-range", "One or more OCR confidence values were normalized; non-finite values were removed and out-of-range values were clamped.", true));
        }
        if (discardedHierarchyId) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Ocr,
                "ocr-hierarchy-id-limit", "One or more OCR hierarchy identifiers exceeded 256 characters and were discarded.", true));
        }

        IReadOnlyList<OcrDiagnostic> returnedDiagnostics = result.Diagnostics ?? Array.Empty<OcrDiagnostic>();
        int diagnosticLimit = Math.Min(returnedDiagnostics.Count, options.MaxProviderDiagnosticsPerCandidate);
        var providerDiagnostics = new List<OcrDiagnostic>(diagnosticLimit);
        int remainingDiagnosticCharacters = options.MaxProviderDiagnosticCharactersPerCandidate;
        int remainingDiagnosticAttributes = options.MaxProviderDiagnosticAttributesPerCandidate;
        int remainingDiagnosticAttributeCharacters = options.MaxProviderDiagnosticAttributeCharactersPerCandidate;
        bool truncatedDiagnosticCharacters = false;
        bool truncatedDiagnosticAttributes = false;
        bool truncatedDiagnosticAttributeCharacters = false;
        for (int index = 0; index < diagnosticLimit; index++) {
            OcrDiagnostic? diagnostic = returnedDiagnostics[index];
            if (diagnostic == null) continue;
            providerDiagnostics.Add(SanitizeProviderDiagnostic(
                diagnostic,
                engineId,
                ref remainingDiagnosticCharacters,
                ref remainingDiagnosticAttributes,
                ref remainingDiagnosticAttributeCharacters,
                ref truncatedDiagnosticCharacters,
                ref truncatedDiagnosticAttributes,
                ref truncatedDiagnosticAttributeCharacters));
        }
        result.Diagnostics = providerDiagnostics.ToArray();
        if (returnedDiagnostics.Count > options.MaxProviderDiagnosticsPerCandidate) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-provider-diagnostic-limit", "OCR provider diagnostics were truncated at MaxProviderDiagnosticsPerCandidate.", true));
        }
        if (truncatedDiagnosticCharacters) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-provider-diagnostic-text-limit", "OCR provider diagnostic text was truncated at MaxProviderDiagnosticCharactersPerCandidate.", true));
        }
        if (truncatedDiagnosticAttributes) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-provider-diagnostic-attribute-limit", "OCR provider diagnostic attributes were truncated at MaxProviderDiagnosticAttributesPerCandidate.", true));
        }
        if (truncatedDiagnosticAttributeCharacters) {
            executionDiagnostics.Add(BuildDiagnostic(candidate, null, engineId, OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                "ocr-provider-diagnostic-attribute-text-limit", "OCR provider diagnostic attribute text was truncated at MaxProviderDiagnosticAttributeCharactersPerCandidate.", true));
        }
    }

    private static OcrDiagnostic SanitizeProviderDiagnostic(
        OcrDiagnostic diagnostic,
        string engineId,
        ref int remainingDiagnosticCharacters,
        ref int remainingDiagnosticAttributes,
        ref int remainingDiagnosticAttributeCharacters,
        ref bool truncatedDiagnosticCharacters,
        ref bool truncatedDiagnosticAttributes,
        ref bool truncatedDiagnosticAttributeCharacters) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        if (diagnostic.Attributes != null) {
            int inspectedAttributes = 0;
            foreach (KeyValuePair<string, string> attribute in diagnostic.Attributes) {
                if (remainingDiagnosticAttributes <= 0) {
                    truncatedDiagnosticAttributes = true;
                    break;
                }
                if (remainingDiagnosticAttributeCharacters <= 0) {
                    truncatedDiagnosticAttributeCharacters = true;
                    break;
                }

                inspectedAttributes++;
                string key = ConsumeRequiredText(attribute.Key, ref remainingDiagnosticAttributeCharacters, ref truncatedDiagnosticAttributeCharacters);
                string value = ConsumeRequiredText(attribute.Value, ref remainingDiagnosticAttributeCharacters, ref truncatedDiagnosticAttributeCharacters);
                remainingDiagnosticAttributes--;
                if (!attributes.ContainsKey(key)) attributes.Add(key, value);
            }
            if (diagnostic.Attributes.Count > inspectedAttributes && remainingDiagnosticAttributes <= 0) {
                truncatedDiagnosticAttributes = true;
            }
        }

        string code = ConsumeRequiredText(diagnostic.Code, ref remainingDiagnosticCharacters, ref truncatedDiagnosticCharacters);
        string message = ConsumeRequiredText(diagnostic.Message, ref remainingDiagnosticCharacters, ref truncatedDiagnosticCharacters);
        string? source = ConsumeOptionalText(
            diagnostic.Source,
            ref remainingDiagnosticCharacters,
            ref truncatedDiagnosticCharacters);
        if (source == null) {
            source = ConsumeOptionalText(engineId, ref remainingDiagnosticCharacters, ref truncatedDiagnosticCharacters);
        }
        return new OcrDiagnostic {
            Severity = diagnostic.Severity,
            Code = code,
            Message = message,
            Source = source,
            IsRecoverable = diagnostic.IsRecoverable,
            Attributes = attributes
        };
    }

    private static OfficeDocumentDiagnostic MapProviderDiagnostic(
        OcrDiagnostic diagnostic,
        string engineId,
        OfficeDocumentOcrCandidate candidate) => new OfficeDocumentDiagnostic {
            Severity = diagnostic.Severity switch {
                OcrDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
                OcrDiagnosticSeverity.Warning => OfficeDocumentDiagnosticSeverity.Warning,
                _ => OfficeDocumentDiagnosticSeverity.Information
            },
            Category = OfficeDocumentDiagnosticCategory.Ocr,
            Code = diagnostic.Code ?? string.Empty,
            Message = diagnostic.Message ?? string.Empty,
            Source = string.IsNullOrWhiteSpace(diagnostic.Source) ? engineId : diagnostic.Source,
            IsRecoverable = diagnostic.IsRecoverable,
            Location = candidate.Location,
            Attributes = diagnostic.Attributes == null
                ? new Dictionary<string, string>(StringComparer.Ordinal)
                : diagnostic.Attributes.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal)
        };

    private static OcrRegion? ToOcrRegion(OfficeDocumentRegion? region) => region == null
        ? null
        : new OcrRegion {
            X = region.X,
            Y = region.Y,
            Width = region.Width,
            Height = region.Height
        };

    private static string? NormalizeHierarchyId(string? value, ref bool discarded) {
        if (string.IsNullOrEmpty(value)) return null;
        string raw = value!;
        if (raw.Length > 256) {
            discarded = true;
            return null;
        }
        string normalized = raw.Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static string ConsumeRequiredText(string? value, ref int remainingCharacters, ref bool truncated) {
        return ConsumeOptionalText(value, ref remainingCharacters, ref truncated) ?? string.Empty;
    }

    private static string? ConsumeOptionalText(string? value, ref int remainingCharacters, ref bool truncated) {
        if (string.IsNullOrEmpty(value)) return null;
        string raw = value!;
        if (remainingCharacters <= 0) {
            truncated = true;
            return null;
        }

        int consumedCharacters = Math.Min(raw.Length, remainingCharacters);
        if (raw.Length > remainingCharacters) truncated = true;
        string bounded = consumedCharacters == raw.Length
            ? raw
            : TruncateText(raw, consumedCharacters);
        remainingCharacters -= consumedCharacters;
        string normalized = bounded.Trim();
        return normalized.Length == 0 ? null : normalized;
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
