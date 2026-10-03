using System.Diagnostics;
using OfficeIMO.Ocr;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    // One budget belongs to one public operation, never to an individual attachment.
    private sealed class ExecutionBudget {
        private readonly Stopwatch _elapsed = Stopwatch.StartNew();
        private readonly ExecutionOptionsSnapshot _options;

        internal ExecutionBudget(ExecutionOptionsSnapshot options) {
            _options = options;
            RemainingCharacters = options.MaxTotalRecognizedCharacters;
            RemainingSpans = options.MaxTotalSpans;
            RemainingSpanCharacters = options.MaxTotalSpanCharacters;
        }

        internal int SelectedCandidates { get; set; }
        internal long ReservedInputBytes { get; set; }
        internal int RemainingCharacters { get; private set; }
        internal int RemainingSpans { get; private set; }
        internal int RemainingSpanCharacters { get; private set; }
        internal TimeSpan RemainingTime => _options.TotalTimeout - _elapsed.Elapsed;
        internal bool CanRecognize => RemainingCharacters > 0 && RemainingTime > TimeSpan.Zero;

        internal void Consume(OcrResult result) {
            RemainingCharacters -= result.Text.Length;
            RemainingSpans -= result.Spans.Count;
            foreach (OcrTextSpan span in result.Spans) {
                RemainingSpanCharacters -= span.Text.Length + (span.Language?.Length ?? 0)
                    + (span.BlockId?.Length ?? 0) + (span.ParagraphId?.Length ?? 0) + (span.LineId?.Length ?? 0);
            }
        }
    }
}
