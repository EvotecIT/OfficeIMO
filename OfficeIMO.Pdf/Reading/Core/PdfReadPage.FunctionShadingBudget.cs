using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal sealed partial class PageContentBudget {
        private long _functionShadingWork;
        private long _functionShadingPixels;
        internal double FunctionShadingScale { get; set; } = 1D;
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
