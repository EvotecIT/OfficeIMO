using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static void ApplyTableBorders(ExcelCell target, OfficeBorderBox borders) {
        Apply(borders.Left, ExcelCellBorders.Left); Apply(borders.Top, ExcelCellBorders.Top);
        Apply(borders.Right, ExcelCellBorders.Right); Apply(borders.Bottom, ExcelCellBorders.Bottom);

        void Apply(OfficeBorderSide? side, ExcelCellBorders edge) {
            if (side is not { } value) return;
            target.SetBorder(edge, value.IsVisible ? ExcelBorderWidth(value.Width) : ExcelBorderStyle.None,
                value.IsVisible ? value.Color.ToRgbHex() : null);
        }
    }

    private static ExcelBorderStyle ExcelBorderWidth(double points) => points < 0.75d ? ExcelBorderStyle.Hair
        : points < 1.5d ? ExcelBorderStyle.Thin : points < 2.5d ? ExcelBorderStyle.Medium : ExcelBorderStyle.Thick;

    private static IEnumerable<double> TableBorderWidths(IWorkTable table, CancellationToken token) {
        if (!table.HasBorderDeclarations) yield break;
        for (int row = 1; row <= table.RowCount; row++) {
            token.ThrowIfCancellationRequested();
            for (int column = 1; column <= table.ColumnCount; column++) {
                var borders = table.GetBorders(row, column);
                if (borders.Left is { IsVisible: true } left) yield return left.Width;
                if (borders.Top is { IsVisible: true } top) yield return top.Width;
                if (borders.Right is { IsVisible: true } right) yield return right.Width;
                if (borders.Bottom is { IsVisible: true } bottom) yield return bottom.Width;
            }
        }
    }
}
