using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

internal static partial class WordDocumentImageRenderer {
    private static bool AddChart(
        WordChart chart,
        WordImageFlowContext context,
        List<OfficeImageExportDiagnostic> diagnostics) {
        if (!chart.TryGetSnapshot(out WordChartSnapshot snapshot)) {
            if (context.IsTargetPage) {
                AddDiagnostic(
                    diagnostics,
                    WordImageExportDiagnosticCodes.UnsupportedChart,
                    "Skipped a Word chart because its cached chart data could not be projected through the shared Drawing chart renderer.",
                    "Word chart");
            }
            return false;
        }

        double width = Math.Min(Math.Max(1D, snapshot.WidthPoints), context.ContentWidth);
        double height = Math.Max(1D, snapshot.HeightPoints);
        if (snapshot.WidthPoints > 0D && width < snapshot.WidthPoints) {
            height *= width / snapshot.WidthPoints;
        }
        height = Math.Min(height, context.ContentHeight);
        if (!EnsureVerticalSpace(context, height, diagnostics)) {
            return false;
        }

        if (context.IsTargetPage) {
            try {
                OfficeChartSnapshot drawingSnapshot = CreateOfficeChartSnapshot(snapshot, width, height);
                OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(
                    drawingSnapshot,
                    useMinimumCanvas: false);
                context.Drawing.AddDrawing(chartDrawing, context.Left, context.Y);
            } catch (Exception exception) when (
                exception is ArgumentException
                || exception is InvalidOperationException
                || exception is NotSupportedException
                || exception is OverflowException) {
                AddDiagnostic(
                    diagnostics,
                    WordImageExportDiagnosticCodes.UnsupportedChart,
                    "Skipped a Word chart because the shared Drawing chart renderer rejected its cached data: " + exception.Message,
                    "Word chart");
                return false;
            }
        }

        context.Y += height + ParagraphGapPoints;
        return true;
    }

    private static OfficeChartSnapshot CreateOfficeChartSnapshot(
        WordChartSnapshot snapshot,
        double width,
        double height) {
        var series = new List<OfficeChartSeries>(snapshot.Data.Series.Count);
        foreach (WordChartSeries item in snapshot.Data.Series) {
            series.Add(new OfficeChartSeries(
                item.Name,
                item.Values,
                item.XValues,
                color: item.Color,
                pointColors: item.PointColors,
                showMarkers: true,
                renderKind: MapChartKind(snapshot.ChartKind)));
        }

        return new OfficeChartSnapshot(
            snapshot.Name,
            snapshot.Title,
            MapChartKind(snapshot.ChartKind),
            new OfficeChartData(snapshot.Data.Categories, series),
            width,
            height);
    }

    private static OfficeChartKind MapChartKind(WordChartSnapshotKind kind) => kind switch {
        WordChartSnapshotKind.ClusteredColumn => OfficeChartKind.ColumnClustered,
        WordChartSnapshotKind.StackedColumn => OfficeChartKind.ColumnStacked,
        WordChartSnapshotKind.StackedColumn100 => OfficeChartKind.ColumnStacked100,
        WordChartSnapshotKind.ClusteredBar => OfficeChartKind.BarClustered,
        WordChartSnapshotKind.StackedBar => OfficeChartKind.BarStacked,
        WordChartSnapshotKind.StackedBar100 => OfficeChartKind.BarStacked100,
        WordChartSnapshotKind.Line => OfficeChartKind.Line,
        WordChartSnapshotKind.StackedLine => OfficeChartKind.LineStacked,
        WordChartSnapshotKind.StackedLine100 => OfficeChartKind.LineStacked100,
        WordChartSnapshotKind.Area => OfficeChartKind.Area,
        WordChartSnapshotKind.StackedArea => OfficeChartKind.AreaStacked,
        WordChartSnapshotKind.StackedArea100 => OfficeChartKind.AreaStacked100,
        WordChartSnapshotKind.Radar => OfficeChartKind.Radar,
        WordChartSnapshotKind.Scatter => OfficeChartKind.Scatter,
        WordChartSnapshotKind.Pie => OfficeChartKind.Pie,
        WordChartSnapshotKind.Doughnut => OfficeChartKind.Doughnut,
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported Word chart snapshot kind.")
    };
}
