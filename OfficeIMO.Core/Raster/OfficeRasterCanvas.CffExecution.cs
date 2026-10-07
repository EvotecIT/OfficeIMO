using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // A retained run needs its own execution state, while contour measurement and
    // painting remain charged to the document's cumulative operation allowance.
    internal void ShareCffOperationBudget(OfficeCffOperationBudget budget) =>
        _cffOperationBudget = budget?.CreateExecutionScope() ?? throw new ArgumentNullException(nameof(budget));

    internal void ShareCffOperationBudget(OfficeRasterCanvas canvas) => ShareCffOperationBudget(canvas._cffOperationBudget);

    // Reset a transformed run's measurement sequence to match its fresh paint
    // canvas, then restore the surrounding direct-text sequence without refunding work.
    internal IDisposable PushCffExecutionScope() {
        var scope = new CffExecutionScope(this, _cffOperationBudget);
        _cffOperationBudget = _cffOperationBudget.CreateExecutionScope();
        return scope;
    }

    private sealed class CffExecutionScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly OfficeCffOperationBudget _previous;
        internal CffExecutionScope(OfficeRasterCanvas canvas, OfficeCffOperationBudget previous) { _canvas = canvas; _previous = previous; }
        public void Dispose() { if (_canvas != null) { _canvas._cffOperationBudget = _previous; _canvas = null; } }
    }
}
