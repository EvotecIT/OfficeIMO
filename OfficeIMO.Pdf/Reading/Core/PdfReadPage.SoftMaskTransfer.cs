namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool TryCreateSoftMaskTransferTable(PdfObject transferObject, PageContentBudget? budget, out byte[]? table) {
        table = null;
        // Invoked-resource discovery needs graphics-state metadata only. Every evaluator/drawing
        // caller supplies its shared operation budget; discovery must never execute a transfer.
        if (budget == null) return true;
        budget.CancellationToken.ThrowIfCancellationRequested();
        PdfObject key = ResolveEffectObject(transferObject) ?? transferObject;
        if (budget.SoftMaskTransferTables.TryGetValue(key, out table)) return true;
        if (!PdfColorSpaceFunctionResolver.TryCreateFunction(transferObject, 1, 1, _objects,
                _limits.MaxDecodedStreamBytes, budget.ColorFunctionResolutionContext, out PdfColorFunction function)) return false;
        var samples = new byte[256];
        var input = new double[1];
        var output = new double[1];
        for (int index = 0; index < samples.Length; index++) {
            budget.ChargeFunctionShadingWork(function.EvaluationCost);
            input[0] = index / 255D;
            if (!function.TryEvaluate(input, output) || double.IsNaN(output[0])) return false;
            samples[index] = (byte)Math.Round(Math.Max(0D, Math.Min(1D, output[0])) * 255D);
        }
        budget.CancellationToken.ThrowIfCancellationRequested();
        budget.SoftMaskTransferTables.Add(key, samples);
        table = samples;
        return true;
    }

    internal sealed partial class PageContentBudget {
        internal Dictionary<PdfObject, byte[]> SoftMaskTransferTables { get; } = new();
    }
}
