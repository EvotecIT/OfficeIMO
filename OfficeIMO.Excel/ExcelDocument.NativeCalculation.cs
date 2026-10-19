namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Applies the ordinary calculation policy before native serialization and carries the unprojected-formula cache policy to the binary owner.</summary>
    internal ExcelNativeCalculationPolicy PrepareNativeCalculationForSave(ExcelSaveOptions? options) {
        bool clear = ShouldClearCachedFormulaResultsBeforeSave(options) && !ShouldEvaluateFormulasBeforeSave(options);
        bool dirty = clear || ShouldMarkFormulasDirtyBeforeSave(options) || ShouldForceFullCalculationOnOpen(options);
        ApplyCalculationPolicyBeforeSave(options);
        if (clear) ConfigureFullCalculationOnOpen();
        return new ExcelNativeCalculationPolicy(clear, dirty);
    }
}

internal readonly record struct ExcelNativeCalculationPolicy(bool ClearCachedResults, bool RecalculateOnOpen);
