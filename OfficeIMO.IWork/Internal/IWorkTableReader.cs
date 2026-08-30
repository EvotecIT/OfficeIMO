namespace OfficeIMO.IWork.Internal;

/// <summary>Shared table projection entry point used by Pages, Numbers, and Keynote graphs.</summary>
internal static class IWorkTableReader {
    internal static IWorkTable? Read(IWorkSourceDocument source, IWorkArchiveRecord tableRecord,
        IWorkProjectionBudget projectionBudget, List<IWorkDiagnostic> diagnostics,
        ref int materializedCellCount,
        ref bool supportsEditableReconstruction) =>
        IWorkNumbersReader.ReadTableInfo(source, tableRecord, projectionBudget, diagnostics,
            ref materializedCellCount, ref supportsEditableReconstruction);
}
