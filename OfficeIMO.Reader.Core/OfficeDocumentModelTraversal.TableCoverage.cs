using System;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    // Equal payloads can carry different observations of source coverage. Preserve the
    // most conservative observation without changing either caller-owned projection.
    private static ReaderTable MergeTableProjection(ReaderTable primary, ReaderTable alias) {
        ReaderTable merged = WithLocationFallback(primary, alias.Location ?? new ReaderLocation(), alias.Location?.TableIndex);
        merged.Truncated = primary.Truncated || alias.Truncated;
        merged.TotalRowCount = Math.Max(primary.TotalRowCount, alias.TotalRowCount);
        merged.Diagnostics = MergeTableDiagnostics(primary.Diagnostics, alias.Diagnostics);
        if (merged.ColumnProfiles == null || merged.ColumnProfiles.Count == 0) merged.ColumnProfiles = alias.ColumnProfiles;
        return merged;
    }

    private static ReaderTableDiagnostics? MergeTableDiagnostics(ReaderTableDiagnostics? primary, ReaderTableDiagnostics? alias) {
        if (primary == null) return alias;
        if (alias == null || ReferenceEquals(primary, alias)) return primary;
        ReaderTableDiagnostics geometry = primary.HasGeometry ? primary : alias;
        return new ReaderTableDiagnostics {
            Confidence = Math.Min(primary.Confidence, alias.Confidence),
            SchemaConfidence = Math.Min(primary.SchemaConfidence, alias.SchemaConfidence),
            CellCompleteness = Math.Min(primary.CellCompleteness, alias.CellCompleteness),
            ColumnGeometryConfidence = Math.Min(primary.ColumnGeometryConfidence, alias.ColumnGeometryConfidence),
            SourceRowCount = Math.Max(primary.SourceRowCount, alias.SourceRowCount),
            ExpectedCellCount = Math.Max(primary.ExpectedCellCount, alias.ExpectedCellCount),
            FilledCellCount = Math.Min(primary.FilledCellCount, alias.FilledCellCount),
            MissingCellCount = Math.Max(primary.MissingCellCount, alias.MissingCellCount),
            HasGeometry = geometry.HasGeometry,
            XStart = geometry.XStart,
            XEnd = geometry.XEnd,
            YTop = geometry.YTop,
            YBottom = geometry.YBottom,
            Width = geometry.Width,
            Height = geometry.Height
        };
    }
}
