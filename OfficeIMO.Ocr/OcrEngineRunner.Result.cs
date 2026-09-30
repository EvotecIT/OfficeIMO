using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Ocr;

public static partial class OcrEngineRunner {
    private static OcrResult CaptureProviderResult(OcrResult result) {
        if (result == null) return null!; // The caller classifies a null result separately.
        return new OcrResult {
            Text = result.Text, Confidence = result.Confidence, Language = result.Language,
            Provider = result.Provider, Model = result.Model,
            Orientation = result.Orientation == null ? null : new OcrOrientationResult {
                ClockwiseRotationDegrees = result.Orientation.ClockwiseRotationDegrees,
                Confidence = result.Orientation.Confidence, Script = result.Orientation.Script
            },
            Spans = (result.Spans ?? Array.Empty<OcrTextSpan>()).Select(CaptureSpan).ToArray(),
            Diagnostics = (result.Diagnostics ?? Array.Empty<OcrDiagnostic>()).Select(CaptureDiagnostic).ToArray()
        };
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

    private static OcrDiagnostic CaptureDiagnostic(OcrDiagnostic diagnostic) {
        if (diagnostic == null) return null!;
        return new OcrDiagnostic {
            Severity = diagnostic.Severity, Code = diagnostic.Code, Message = diagnostic.Message,
            Source = diagnostic.Source, IsRecoverable = diagnostic.IsRecoverable,
            Attributes = diagnostic.Attributes == null ? new Dictionary<string, string>(StringComparer.Ordinal)
                : diagnostic.Attributes.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal)
        };
    }
}
