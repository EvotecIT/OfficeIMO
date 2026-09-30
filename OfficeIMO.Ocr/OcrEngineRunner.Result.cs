using System;
using System.Collections.Generic;

namespace OfficeIMO.Ocr;

public static partial class OcrEngineRunner {
    private static OcrResult CaptureProviderResult(OcrResult result, OcrResultCaptureLimits limits, Action checkDeadline) {
        checkDeadline();
        if (result == null) return null!; // The caller classifies a null result separately.
        var captured = new OcrResult {
            Text = result.Text, Confidence = result.Confidence, Language = result.Language,
            Provider = result.Provider, Model = result.Model,
            Orientation = result.Orientation == null ? null : new OcrOrientationResult {
                ClockwiseRotationDegrees = result.Orientation.ClockwiseRotationDegrees,
                Confidence = result.Orientation.Confidence, Script = result.Orientation.Script
            }
        };
        IReadOnlyList<OcrDiagnostic> diagnostics = result.Diagnostics ?? Array.Empty<OcrDiagnostic>();
        int diagnosticCount = diagnostics.Count;
        var retainedDiagnostics = new List<OcrDiagnostic>();
        int remainingAttributes = limits.MaxDiagnosticAttributes;
        // Inspect terminal severity outside the retained prefix, without copying discarded attributes.
        for (int index = 0; index < diagnosticCount; index++) {
            checkDeadline();
            OcrDiagnostic diagnostic = diagnostics[index];
            checkDeadline();
            if (diagnostic != null && diagnostic.Severity == OcrDiagnosticSeverity.Error && !diagnostic.IsRecoverable) {
                captured.Diagnostics = new[] { new OcrDiagnostic { Severity = OcrDiagnosticSeverity.Error, IsRecoverable = false } };
                return captured;
            }
            if (index < limits.MaxDiagnostics) retainedDiagnostics.Add(CaptureDiagnostic(diagnostic, ref remainingAttributes, checkDeadline));
        }
        captured.Diagnostics = retainedDiagnostics.ToArray();
        captured.OmittedDiagnosticCount = diagnosticCount - retainedDiagnostics.Count;
        IReadOnlyList<OcrTextSpan> spans = result.Spans ?? Array.Empty<OcrTextSpan>();
        int spanCount = spans.Count;
        int spanLimit = Math.Min(spanCount, limits.MaxSpans);
        var retainedSpans = new List<OcrTextSpan>();
        for (int index = 0; index < spanLimit; index++) {
            checkDeadline();
            retainedSpans.Add(CaptureSpan(spans[index]));
            checkDeadline();
        }
        captured.Spans = retainedSpans.ToArray();
        captured.OmittedSpanCount = spanCount - spanLimit;
        checkDeadline();
        return captured;
    }

    private static OcrTextSpan CaptureSpan(OcrTextSpan span) {
        if (span == null) return null!;
        return new OcrTextSpan {
            Sequence = span.Sequence, Level = span.Level, Text = span.Text,
            Confidence = span.Confidence, Language = span.Language, PageNumber = span.PageNumber,
            BlockId = span.BlockId, ParagraphId = span.ParagraphId, LineId = span.LineId,
            CoordinateUnit = span.CoordinateUnit,
            Region = span.Region == null ? null : new OcrRegion {
                X = span.Region.X, Y = span.Region.Y, Width = span.Region.Width, Height = span.Region.Height
            }
        };
    }

    private static OcrDiagnostic CaptureDiagnostic(OcrDiagnostic? diagnostic, ref int remainingAttributes, Action checkDeadline) {
        if (diagnostic == null) return null!;
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        int sourceCount = diagnostic.Attributes?.Count ?? 0;
        if (sourceCount > 0 && remainingAttributes > 0) {
            using var enumerator = diagnostic.Attributes!.GetEnumerator();
            while (remainingAttributes > 0 && attributes.Count < sourceCount) {
                checkDeadline();
                if (!enumerator.MoveNext()) break;
                checkDeadline();
                var pair = enumerator.Current;
                attributes.Add(pair.Key, pair.Value);
                remainingAttributes--;
            }
        }
        return new OcrDiagnostic {
            Severity = diagnostic.Severity, Code = diagnostic.Code, Message = diagnostic.Message,
            Source = diagnostic.Source, IsRecoverable = diagnostic.IsRecoverable,
            Attributes = attributes, OmittedAttributeCount = sourceCount - attributes.Count
        };
    }

    private sealed class CaptureDeadlineException : Exception {
        internal CaptureDeadlineException(string engineId, TimeSpan timeout) { EngineId = engineId; Timeout = timeout; }
        internal string EngineId { get; }
        internal TimeSpan Timeout { get; }
    }
}
