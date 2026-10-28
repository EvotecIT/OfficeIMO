using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Adds an editable native chart from typed category data and numeric series.</summary>
    public PowerPointChart AddChart<T, TCategory>(OfficeChartType type, IEnumerable<T> data,
        Func<T, TCategory> category, params OfficeSeries[] series) =>
        AddChart(type, OfficeTypedChartData.Create(data, category, series));

    /// <summary>Adds an editable native chart using property names, including names supplied with nameof.</summary>
    public PowerPointChart AddChart<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        OfficeChartType type, IEnumerable<T> data, string category, params string[] series) =>
        AddChart(type, OfficeTypedChartData.Create(data, category, series));
}
