using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal sealed partial class PageContentBudget {
        private long _functionShadingWork;
        private long _functionShadingPixels;
        private double _functionShadingScale = 1D;

        /// <summary>The scale function shadings are sampled for. Reading it marks the page as drawn for that scale.</summary>
        internal double FunctionShadingScale {
            get {
                UsesFunctionShadingScale = true;
                return _functionShadingScale;
            }
            set => _functionShadingScale = value;
        }

        /// <summary>Whether the content read <see cref="FunctionShadingScale"/>, so a drawing for another scale differs.</summary>
        internal bool UsesFunctionShadingScale { get; private set; }
        internal long MaximumFunctionPixels { get; set; } = OfficeImageExportOptions.DefaultMaximumRasterPixels;

        internal bool ChargeFunctionShadingWork(int work) {
            CancellationToken.ThrowIfCancellationRequested();
            _functionShadingWork += Math.Max(1, work);
            if (_functionShadingWork > _page._limits.MaxFunctionShadingEvaluationWork)
                throw PdfReadLimitException.Create(PdfReadLimitKind.FunctionShadingEvaluationWork,
                    _page._limits.MaxFunctionShadingEvaluationWork, _functionShadingWork);
            return true;
        }

        internal void ChargeFunctionShadingPixels(long pixels) {
            CancellationToken.ThrowIfCancellationRequested();
            if (pixels > MaximumFunctionPixels)
                throw PdfReadLimitException.Create(PdfReadLimitKind.FunctionShadingPixels, MaximumFunctionPixels, pixels);
            long limit = _page._limits.MaxFunctionShadingPixels;
            if (pixels < 0 || pixels > limit - _functionShadingPixels)
                throw PdfReadLimitException.Create(PdfReadLimitKind.FunctionShadingPixels, limit,
                    pixels > long.MaxValue - _functionShadingPixels ? long.MaxValue : _functionShadingPixels + pixels);
            _functionShadingPixels += pixels;
        }
    }
}
