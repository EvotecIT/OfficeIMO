using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static void ApplyTableBorders(WordTableCell target, OfficeBorderBox borders, bool clearMissing) {
        var result = target.Borders;
        if (clearMissing) {
            result.LeftStyle = result.TopStyle = result.RightStyle = result.BottomStyle = WordBorderStyle.Nil;
        }
        if (borders.Left is { } left) {
            result.LeftStyle = left.IsVisible ? WordBorderStyle.Single : WordBorderStyle.Nil;
            result.LeftColorHex = left.IsVisible ? left.Color.ToRgbHex() : null;
            result.LeftSize = left.IsVisible ? BorderSize(left.Width) : null;
        }
        if (borders.Top is { } top) {
            result.TopStyle = top.IsVisible ? WordBorderStyle.Single : WordBorderStyle.Nil;
            result.TopColorHex = top.IsVisible ? top.Color.ToRgbHex() : null;
            result.TopSize = top.IsVisible ? BorderSize(top.Width) : null;
        }
        if (borders.Right is { } right) {
            result.RightStyle = right.IsVisible ? WordBorderStyle.Single : WordBorderStyle.Nil;
            result.RightColorHex = right.IsVisible ? right.Color.ToRgbHex() : null;
            result.RightSize = right.IsVisible ? BorderSize(right.Width) : null;
        }
        if (borders.Bottom is { } bottom) {
            result.BottomStyle = bottom.IsVisible ? WordBorderStyle.Single : WordBorderStyle.Nil;
            result.BottomColorHex = bottom.IsVisible ? bottom.Color.ToRgbHex() : null;
            result.BottomSize = bottom.IsVisible ? BorderSize(bottom.Width) : null;
        }
    }

    private static uint BorderSize(double points) => checked((uint)Math.Round(points * 8d, MidpointRounding.AwayFromZero));

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
