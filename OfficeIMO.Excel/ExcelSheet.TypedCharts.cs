using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an editable native chart from typed category data and numeric series.</summary>
    public ExcelChart AddChart<T, TCategory>(OfficeChartType type, IEnumerable<T> data,
        Func<T, TCategory> category, params OfficeSeries[] series) =>
        AddChart(type, OfficeTypedChartData.Create(data, category, series), 1, 1);

    /// <summary>Adds an editable native chart using property names, including names supplied with nameof.</summary>
    public ExcelChart AddChart<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        OfficeChartType type, IEnumerable<T> data, string category, params string[] series) =>
        AddChart(type, OfficeTypedChartData.Create(data, category, series), 1, 1);
}
