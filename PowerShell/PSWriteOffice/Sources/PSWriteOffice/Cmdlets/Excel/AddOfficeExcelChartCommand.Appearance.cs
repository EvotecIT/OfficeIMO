using System;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class AddOfficeExcelChartCommand {
    private (int Row, int Column, int Width, int Height) ResolvePlacement() {
        var (row, column) = ExcelA1.ParseCellRef(Address);
        if (row < 1 || row > ExcelA1.MaxRows || column < 1 || column > ExcelA1.MaxColumns) {
            throw new PSArgumentException("Address must be one worksheet cell such as F2.", nameof(Address));
        }
        return (row, column, PixelSize(Width, nameof(Width)), PixelSize(Height, nameof(Height)));
    }

    /// <summary>Rounds at the native pixel boundary and rejects unusable sizes before any worksheet writes.</summary>
    private static int PixelSize(OfficeLength length, string parameter) {
        double pixels = Math.Round(length.ToPixels(), MidpointRounding.AwayFromZero);
        if (pixels < 1 || pixels > int.MaxValue) throw new PSArgumentException("Chart dimensions must round to a positive Int32 pixel size.", parameter);
        return checked((int)pixels);
    }

    private OfficeChartType ResolveChartKind() {
        if (!Enum.IsDefined(typeof(ExcelDashboardChartPreset), Preset)) throw new PSArgumentException("Unknown chart preset.", nameof(Preset));
        var kind = MyInvocation.BoundParameters.ContainsKey(nameof(Type)) ? Type : Preset switch {
            ExcelDashboardChartPreset.Trend => OfficeChartType.Line,
            ExcelDashboardChartPreset.Contribution => OfficeChartType.Pie,
            ExcelDashboardChartPreset.CompactComparison => OfficeChartType.BarClustered,
            _ => Type
        };
        ValidateChartType(kind);
        return kind;
    }

    private static void ValidateChartType(OfficeChartType kind) {
        if (!Enum.IsDefined(typeof(OfficeChartType), kind) || kind == OfficeChartType.Bubble) {
            throw new PSArgumentException($"Chart type '{kind}' is not supported by this command. Bubble charts require explicit size data.", nameof(Type));
        }
    }

    private static void ValidateChartData(OfficeChartData data, OfficeChartType kind) {
        if ((kind is OfficeChartType.Pie or OfficeChartType.Doughnut) && data.Series.Count != 1) {
            throw new PSArgumentException("Pie and Doughnut charts require exactly one value column.", nameof(Value));
        }
        if (kind == OfficeChartType.Scatter && data.Categories.Any(category =>
            !double.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ||
            double.IsNaN(number) || double.IsInfinity(number))) {
            throw new PSArgumentException("Scatter charts require finite numeric categories.", nameof(Category));
        }
    }
}
