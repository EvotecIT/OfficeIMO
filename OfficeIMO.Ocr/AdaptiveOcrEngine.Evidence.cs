using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Ocr;

public sealed partial class AdaptiveOcrEngine {
    private void ValidateInput(OcrRequest request, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (request.Payload == null || request.Payload.Length == 0 || request.Payload.LongLength > _maximumInputBytes)
            throw new ArgumentException("The raster must be nonempty and within the configured input budget.", nameof(request));
    }

    private void CheckTextBudget(OcrResult result) {
        long characters = result.Text?.Length ?? 0;
        foreach (OcrTextSpan span in result.Spans) {
            characters += span?.Text?.Length ?? 0;
            if (characters > _maximumTextCharacters) throw new OcrEngineExecutionException(OcrEngineFailureKind.InvalidResult);
        }
        if (characters > _maximumTextCharacters) throw new OcrEngineExecutionException(OcrEngineFailureKind.InvalidResult);
    }

    private static OcrRequest CopyRequest(OcrRequest request) => new OcrRequest {
        Operation = request.Operation, Payload = (byte[])request.Payload.Clone(), MediaType = request.MediaType,
        FileName = request.FileName, SourceId = request.SourceId, SourceName = request.SourceName,
        CandidateId = request.CandidateId, CandidateKind = request.CandidateKind, PageNumber = request.PageNumber,
        PixelWidth = request.PixelWidth, PixelHeight = request.PixelHeight, Language = request.Language,
        RegionCoordinateUnit = request.RegionCoordinateUnit,
        Region = request.Region == null ? null : new OcrRegion { X = request.Region.X, Y = request.Region.Y,
            Width = request.Region.Width, Height = request.Region.Height },
        ProviderOptions = (request.ProviderOptions ?? new Dictionary<string, string>())
            .ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal)
    };

    private static string Normalize(string? text, CancellationToken token) {
        var builder = new StringBuilder();
        bool space = false;
        string source = (text ?? string.Empty).Normalize(NormalizationForm.FormC);
        for (int index = 0; index < source.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            char value = source[index];
            if (char.IsWhiteSpace(value)) { space = builder.Length > 0; continue; }
            if (space) builder.Append(' ');
            builder.Append(value); space = false;
        }
        return builder.ToString();
    }

    private void AddReviewDiagnostic(AdaptiveOcrResult result) {
        string Value(int number) => number.ToString(CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["review-status"] = result.Review.Status.ToString(), ["attempts"] = Value(result.Attempts.Count), ["selected-attempt"] = Value(result.SelectedAttempt),
            ["words"] = Value(result.Quality.WordCount), ["low-confidence-words"] = Value(result.Quality.LowConfidenceWordCount),
            ["unknown-confidence-words"] = Value(result.Quality.UnknownConfidenceWordCount),
            ["thresholds-met"] = result.Quality.MeetsThresholds ? "true" : "false",
            ["disagreement"] = result.HasDisagreement ? "true" : "false", ["retry-incomplete"] = result.RetryIncomplete ? "true" : "false"
        };
        result.Result.Diagnostics = new[] { new OcrDiagnostic {
            Source = Id, Code = result.Review.Status == OcrReviewStatus.Unassessed ? "adaptive-ocr-unassessed"
                : result.ReviewRecommended ? "adaptive-ocr-review-recommended" : "adaptive-ocr-thresholds-met",
            Severity = result.ReviewRecommended ? OcrDiagnosticSeverity.Warning : OcrDiagnosticSeverity.Info,
            Message = result.Review.Status == OcrReviewStatus.Unassessed
                ? "Confidence checks passed without a completed comparison; this does not establish text correctness or approval."
                : result.ReviewRecommended ? "Adaptive OCR evidence recommends human review." :
                "OCR evidence meets configured checks; this does not establish text correctness or approval.", Attributes = attributes
        } }.Concat(result.Result.Diagnostics).ToArray();
    }
}
