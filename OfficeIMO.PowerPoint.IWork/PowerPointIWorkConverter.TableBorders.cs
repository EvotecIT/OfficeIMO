using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    private static void ApplyTableBorders(PowerPointTableCell target, OfficeBorderBox borders) {
        Apply(borders.Left, PowerPointTableCellBorders.Left);
        Apply(borders.Top, PowerPointTableCellBorders.Top);
        Apply(borders.Right, PowerPointTableCellBorders.Right);
        Apply(borders.Bottom, PowerPointTableCellBorders.Bottom);

        void Apply(OfficeBorderSide? side, PowerPointTableCellBorders edge) {
            if (side is not { } value) return;
            if (value.IsVisible) target.SetBorders(edge, value.Color.ToRgbHex(), global::OfficeIMO.OfficeLength.Points(value.Width));
            else target.HideBorders(edge);
        }
    }

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
