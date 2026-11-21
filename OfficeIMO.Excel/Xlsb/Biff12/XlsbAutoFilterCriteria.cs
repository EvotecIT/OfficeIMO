using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Biff12;

/// <summary>Maps AutoFilter criteria between SpreadsheetML values and their BIFF12 record codes.</summary>
internal static class XlsbAutoFilterCriteria {
    internal const int BrtColorFilter = 168;
    internal const int BrtIconFilter = 169;
    internal const int BrtTop10Filter = 170;
    internal const int BrtDynamicFilter = 171;
    internal const int BrtBeginCustomFilters = 172;
    internal const int BrtEndCustomFilters = 173;
    internal const int BrtCustomFilter = 174;

    /// <summary>BrtCustomFilter vts (2.4.348) for a real number.</summary>
    internal const byte NumberValue = 0x04;

    /// <summary>BrtCustomFilter vts for a string.</summary>
    internal const byte StringValue = 0x06;

    // BrtCustomFilter grbitSgn 1..6 (2.4.348).
    private static readonly FilterOperatorValues[] Operators = {
        FilterOperatorValues.LessThan, FilterOperatorValues.Equal, FilterOperatorValues.LessThanOrEqual,
        FilterOperatorValues.GreaterThan, FilterOperatorValues.NotEqual, FilterOperatorValues.GreaterThanOrEqual
    };

    // BrtDynamicFilter cft (2.4.362): 0..2, then 0x08..0x27 in this order.
    private static readonly DynamicFilterValues[] DynamicTypes = {
        DynamicFilterValues.Null, DynamicFilterValues.AboveAverage, DynamicFilterValues.BelowAverage,
        DynamicFilterValues.Tomorrow, DynamicFilterValues.Today, DynamicFilterValues.Yesterday,
        DynamicFilterValues.NextWeek, DynamicFilterValues.ThisWeek, DynamicFilterValues.LastWeek,
        DynamicFilterValues.NextMonth, DynamicFilterValues.ThisMonth, DynamicFilterValues.LastMonth,
        DynamicFilterValues.NextQuarter, DynamicFilterValues.ThisQuarter, DynamicFilterValues.LastQuarter,
        DynamicFilterValues.NextYear, DynamicFilterValues.ThisYear, DynamicFilterValues.LastYear,
        DynamicFilterValues.YearToDate,
        DynamicFilterValues.Quarter1, DynamicFilterValues.Quarter2, DynamicFilterValues.Quarter3, DynamicFilterValues.Quarter4,
        DynamicFilterValues.January, DynamicFilterValues.February, DynamicFilterValues.March, DynamicFilterValues.April,
        DynamicFilterValues.May, DynamicFilterValues.June, DynamicFilterValues.July, DynamicFilterValues.August,
        DynamicFilterValues.September, DynamicFilterValues.October, DynamicFilterValues.November, DynamicFilterValues.December
    };

    // KPISets (2.5.86) 0..0x10.
    private static readonly IconSetValues[] IconSets = {
        IconSetValues.ThreeArrows, IconSetValues.ThreeArrowsGray, IconSetValues.ThreeFlags,
        IconSetValues.ThreeTrafficLights1, IconSetValues.ThreeTrafficLights2, IconSetValues.ThreeSigns,
        IconSetValues.ThreeSymbols, IconSetValues.ThreeSymbols2, IconSetValues.FourArrows, IconSetValues.FourArrowsGray,
        IconSetValues.FourRedToBlack, IconSetValues.FourRating, IconSetValues.FourTrafficLights, IconSetValues.FiveArrows,
        IconSetValues.FiveArrowsGray, IconSetValues.FiveRating, IconSetValues.FiveQuarters
    };

    internal static bool TryGetOperator(byte code, out FilterOperatorValues value) {
        value = FilterOperatorValues.Equal;
        if (code < 1 || code > Operators.Length) return false;
        value = Operators[code - 1];
        return true;
    }

    internal static byte GetOperatorCode(FilterOperatorValues value) => checked((byte)(Array.IndexOf(Operators, value) + 1));

    internal static bool TryGetDynamicType(uint code, out DynamicFilterValues value) {
        value = DynamicFilterValues.Null;
        long index = code <= 2 ? code : code >= 0x08 ? code - 5 : -1;
        if (index < 0 || index >= DynamicTypes.Length) return false;
        value = DynamicTypes[index];
        return true;
    }

    internal static bool TryGetDynamicCode(DynamicFilterValues value, out uint code) {
        int index = Array.IndexOf(DynamicTypes, value);
        code = index <= 2 ? (uint)Math.Max(index, 0) : checked((uint)(index + 5));
        return index >= 0;
    }

    /// <summary>Dynamic types whose BrtDynamicFilter carries a computed average (xNumValue only).</summary>
    internal static bool IsAverage(uint code) => code is 0x01 or 0x02;

    /// <summary>Dynamic types whose BrtDynamicFilter carries a computed date range (xNumValue and xNumValueMax).</summary>
    internal static bool IsDateRange(uint code) => code >= 0x08 && code <= 0x17;

    internal static bool TryGetIconSet(uint code, out IconSetValues value, out int iconCount) {
        value = IconSetValues.ThreeArrows;
        iconCount = 0;
        if (code >= IconSets.Length) return false;
        value = IconSets[code];
        iconCount = code <= 7 ? 3 : code <= 12 ? 4 : 5;
        return true;
    }

    internal static bool TryGetIconSetCode(IconSetValues value, out uint code, out int iconCount) {
        int index = Array.IndexOf(IconSets, value);
        code = index < 0 ? 0 : (uint)index;
        iconCount = index <= 7 ? 3 : index <= 12 ? 4 : 5;
        return index >= 0;
    }
}
