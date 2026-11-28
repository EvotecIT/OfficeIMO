using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Write;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private LegacyXlsPreservationSnapshot? _legacyXlsPreservationSnapshot;
    private LegacyXlsDeferredPreservation? _legacyXlsDeferredPreservation;

    internal bool HasOwnedLegacyXlsProjection => _legacyXlsDeferredPreservation != null;

    private sealed record LegacyXlsDeferredPreservation(byte[] Projection, LegacyXlsWorkbook Workbook);

    /// <summary>Reuse the original package at the point before the deferred cell changes are replayed.</summary>
    private void CaptureLegacyXlsBeforeCellReplay(byte[] openedProjection) {
        if (_legacyXlsDeferredPreservation is not { } original || !ReferenceEquals(original.Projection, openedProjection)
            || _packagePropertiesDirty || _customDocumentPropertiesDirty) return;
        _legacyXlsPreservationSnapshot = LegacyXlsPreservationSnapshot.Capture(this, original.Workbook, CancellationToken.None);
        Volatile.Write(ref _legacyXlsDeferredPreservation, null);
    }

    /// <summary>Builds the legacy edit baseline only when a save needs native preservation.</summary>
    private LegacyXlsPreservationSnapshot? GetLegacyXlsPreservationSnapshot(CancellationToken token = default) {
        if (Volatile.Read(ref _legacyXlsDeferredPreservation) == null) return _legacyXlsPreservationSnapshot;

        return WithDeferredOpenGate(() => {
            if (_legacyXlsDeferredPreservation is { } original) {
                token.ThrowIfCancellationRequested();
                using ExcelDocument projection = LoadFromByteArray(original.Projection, new ExcelLoadOptions(),
                    filePath: null, captureSourcePackageSecurity: false);
                projection._deferredCreated = true;
                _legacyXlsPreservationSnapshot = LegacyXlsPreservationSnapshot.Capture(projection, original.Workbook, token);
                Volatile.Write(ref _legacyXlsDeferredPreservation, null);
            }
            return _legacyXlsPreservationSnapshot;
        });
    }

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
        return SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && GetLegacyXlsPreservationSnapshot()?.TryCreatePlan(this, out plan) == true;
    }

    private bool CanPreserveNativeLegacyXls(ExcelSaveOptions? options) =>
        SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && GetLegacyXlsPreservationSnapshot()?.TryCreatePlan(this, out _) == true;

    private bool HasLegacyPreservationTransforms(ExcelSaveOptions? options) =>
        options?.SafeRepairDefinedNames == true || options?.SafePreflight == true || options?.ValidateOpenXml == true;

    /// <summary>Prepares native XLS bytes, or writes an eligible fresh direct workbook to the supplied destination.</summary>
    private byte[] BuildNativeLegacyXls(ExcelSaveOptions? options, CancellationToken cancellationToken, out bool direct, out bool preserved, Stream? directDestination = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (EarlyBiffImportReport != null) {
            throw new NotSupportedException("Early BIFF imports do not support legacy binary workbook output.");
        }
        direct = false; preserved = false;
        var calculation = PrepareNativeCalculationForSave(options);
        if (SourceFormat == ExcelFileFormat.Xls && !HasLegacyPreservationTransforms(options) && GetLegacyXlsPreservationSnapshot(cancellationToken)?.TryCreatePlan(this, out var plan, cancellationToken) == true) {
            preserved = true;
            long maximumBytes = Math.Min(GetPackageMaterializationLimit(options) ?? int.MaxValue, int.MaxValue);
            return LegacyXlsWriter.WithinWorkbookSizeLimit(maximumBytes, () => plan!.Write(null, maximumBytes, cancellationToken, calculation));
        }
        EnsureLegacyXlsSaveDoesNotDropImportedContent(options, preserveLinkedVbaProject: true);
        if (TryBuildDirectTabularLegacyXls(options, cancellationToken, out byte[] bytes, directDestination)) { direct = true; return bytes; }
        PrepareWorkbookForSave(options, applyCalculationPolicy: false);
        cancellationToken.ThrowIfCancellationRequested();
        return LegacyXlsWriter.WriteWorkbook(this);
    }
}
