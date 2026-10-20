using OfficeIMO.Excel.LegacyXls.Write;
using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private LegacyXlsPreservationSnapshot? _legacyXlsPreservationSnapshot;

    internal string GetLegacyXlsCorePropertiesSnapshot() {
        var properties = _spreadSheetDocument.PackageProperties;
        // The package XML formatter uses whole seconds; native OLE timestamps retain individual ticks.
        return string.Join("\0", CreateCorePropertiesXml(_spreadSheetDocument),
            properties.Created?.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            properties.Modified?.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            properties.LastPrinted?.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
    }

    internal bool TryGetLegacyXlsPreservationPlan(ExcelSaveOptions? options, out LegacyXlsPreservationPlan? plan) {
        plan = null;
        return SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && _legacyXlsPreservationSnapshot?.TryCreatePlan(this, out plan) == true;
    }

    private bool CanPreserveNativeLegacyXls(ExcelSaveOptions? options) =>
        SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && _legacyXlsPreservationSnapshot?.TryCreatePlan(this, out _) == true;

    private bool HasLegacyPreservationTransforms(ExcelSaveOptions? options) =>
        options?.SafeRepairDefinedNames == true || options?.SafePreflight == true || options?.ValidateOpenXml == true;

    private byte[] BuildNativeLegacyXls(ExcelSaveOptions? options, CancellationToken cancellationToken, out bool direct, out bool preserved) {
        cancellationToken.ThrowIfCancellationRequested();
        direct = false; preserved = false;
        var calculation = PrepareNativeCalculationForSave(options);
        if (SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && _legacyXlsPreservationSnapshot?.TryCreatePlan(this, out var plan, cancellationToken) == true) {
            preserved = true;
            long maximumBytes = Math.Min(GetPackageMaterializationLimit(options) ?? int.MaxValue, int.MaxValue);
            return plan!.Write(null, maximumBytes, cancellationToken, calculation);
        }
        EnsureLegacyXlsSaveDoesNotDropImportedContent(options, preserveLinkedVbaProject: true);
        if (TryBuildDirectTabularLegacyXls(options, cancellationToken, out byte[] bytes)) { direct = true; return bytes; }
        PrepareWorkbookForSave(options, applyCalculationPolicy: false);
        cancellationToken.ThrowIfCancellationRequested();
        return LegacyXlsWriter.WriteWorkbook(this);
    }
}
