namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private string? _intrinsicTextMeasurementBudgetSource;

    /// <summary>Charges repeated intrinsic text work at its actual measurement entry point, including nested and prepared runs.</summary>
    private void ChargeIntrinsicTextMeasurement(string text) {
        if (_intrinsicTextMeasurementBudgetSource == null) return;
        CheckCancellation();
        ChargeLayoutOperations(text.Length, _intrinsicTextMeasurementBudgetSource);
    }

    /// <summary>Bounds a repeated intrinsic traversal without changing ordinary measurement outside its synchronous lifetime.</summary>
    private readonly struct IntrinsicTextMeasurementBudgetScope : IDisposable {
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly string? _previousSource;

        internal IntrinsicTextMeasurementBudgetScope(HtmlRenderLayoutEngine engine, string source) {
            _engine = engine;
            _previousSource = engine._intrinsicTextMeasurementBudgetSource;
            engine._intrinsicTextMeasurementBudgetSource = source;
        }

        public void Dispose() => _engine._intrinsicTextMeasurementBudgetSource = _previousSource;
    }
}
