using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordTable {
    // Model-only state: source and directly authored grids have their own geometry. Grids
    // generated from preferences must continue to follow later edits to those preferences.
    private sealed class DerivedGrid { }

    private bool HasDerivedGrid => _table.Annotation<DerivedGrid>() != null;

    internal static void CopyGridNormalizationState(WordprocessingDocument source, WordprocessingDocument copy) {
        Dictionary<Uri, OpenXmlPart> parts = copy.GetAllParts().GroupBy(part => part.Uri)
            .ToDictionary(group => group.Key, group => group.First());
        foreach (OpenXmlPart part in source.GetAllParts()) {
            if (part.IsRootElementLoaded && part.RootElement is { } original && parts.TryGetValue(part.Uri, out OpenXmlPart? copiedPart) &&
                copiedPart.RootElement is { } cloned) CopyGridNormalizationState(original, cloned);
        }
    }

    internal static void CopyGridNormalizationState(OpenXmlElement source, OpenXmlElement copy) {
        if (source is Table sourceTable && copy is Table copiedTable) Copy(sourceTable, copiedTable);
        using var originals = source.Descendants<Table>().GetEnumerator();
        using var copies = copy.Descendants<Table>().GetEnumerator();
        while (originals.MoveNext() && copies.MoveNext()) Copy(originals.Current, copies.Current);

        static void Copy(Table original, Table cloned) {
            cloned.RemoveAnnotations<DerivedGrid>();
            if (original.Annotation<DerivedGrid>() != null) cloned.AddAnnotation(new DerivedGrid());
        }
    }

    private int TableWidthNativeValue => WordTableWidthNative.NumericValue(_tableProperties?.TableWidth) ?? 0;
    private double TableWidthFraction => WordTableWidthNative.Percentage(_tableProperties?.TableWidth?.Width?.Value) ?? 0;

    /// <summary>Rebuilds only physical grid geometry. Native preferences are never rewritten by inference.</summary>
    private void RefreshTblGridFromColumnWidths() {
        if (_suppressGridRefresh) return;
        _suppressGridRefresh = true;
        try {
            if (Rows.Count == 0) return;
            // Unknown native declarations are preserved rather than treated as Auto.
            if (!WordTableWidthNative.NativeType(_tableProperties?.TableWidth).HasValue) return;
            int count = _table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
            if (count == 0) count = Rows.Max(row => row.Cells.Sum(cell => Math.Max(1, cell.ColumnSpan)));
            if (count <= 0) return;
            List<WordTableWidth?>? fallback = null;
            List<WordTableWidth?>? preferences = null;
            foreach (var row in Rows) {
                var widths = new List<WordTableWidth?>();
                int offset = (int)(row._tableRow.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0);
                for (int i = 0; i < offset && widths.Count < count; i++) widths.Add(null);
                foreach (var cell in row.Cells) {
                    if (cell.HorizontalMerge == WordCellMerge.Continue) continue;
                    WordTableWidth? width = cell.Width;
                    int span = Math.Max(1, cell.ColumnSpan);
                    for (int i = 0; i < span && widths.Count < count; i++) {
                        widths.Add(width?.Kind == WordTableWidthKind.Length
                            ? WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(width.Value.LengthValue!.Value.ToEmus() / span))
                            : width?.Kind == WordTableWidthKind.Percent ? WordTableWidth.Percent(width.Value.PercentValue!.Value / span) : width);
                    }
                }
                while (widths.Count < count) widths.Add(null);
                if (widths.All(width => width.HasValue)) { preferences = widths; break; }
                if (fallback == null && widths.Any(width => width.HasValue)) fallback = widths;
            }
            preferences ??= fallback ?? Enumerable.Repeat<WordTableWidth?>(null, count).ToList();
            int target = Math.Max(0, EstimateTableWidthInDxa());
            var previous = GridColumnWidth;
            bool allPercent = preferences.All(width => width?.Kind == WordTableWidthKind.Percent);
            bool allLength = preferences.All(width => width?.Kind == WordTableWidthKind.Length);
            bool explicitTable = (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Dxa && TableWidthNativeValue > 0)
                || (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Pct && TableWidthFraction > 0);
            double[] resolved = preferences.Select((width, index) => width?.Kind == WordTableWidthKind.Length
                ? (double)width.Value.LengthValue!.Value.ToTwips()
                : width?.Kind == WordTableWidthKind.Percent ? target * width.Value.PercentValue!.Value / 100d
                : index < previous.Count ? previous[index].ToTwips() : target / (double)count).ToArray();
            // Homogeneous preferences retain existing normalization. Mixed preferences are resolved separately.
            bool normalize = allPercent || (allLength && explicitTable);
            double total = resolved.Sum();
            var columns = new List<global::OfficeIMO.OfficeLength>(count);
            int allocated = 0;
            for (int i = 0; i < count; i++) {
                int twips = checked((int)Math.Round(normalize && total > 0 ? target * resolved[i] / total : resolved[i]));
                if (normalize && total > 0 && i == count - 1) twips = Math.Max(0, target - allocated);
                twips = Math.Max(1, twips);
                allocated = checked(allocated + twips);
                columns.Add(global::OfficeIMO.OfficeLength.Emus((long)twips * 635));
            }
            GridColumnWidth = columns;
            _table.RemoveAnnotations<DerivedGrid>();
            _table.AddAnnotation(new DerivedGrid());
        } catch (NotSupportedException) {
            // Unsupported native preferences retain their original declarations and grid.
        } catch (OverflowException) {
            // Preserve existing geometry if an imported preference cannot be resolved in native range.
        } finally { _suppressGridRefresh = false; }
    }
}
