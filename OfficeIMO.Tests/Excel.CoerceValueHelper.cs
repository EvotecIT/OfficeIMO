using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelCoerceValueHelper
{
    [Fact]
    public void Coerce_NumberAndString()
    {
        var (numVal, numType) = CoerceValueHelper.Coerce(123, s => new CellValue(s));
        Assert.Equal(CellValues.Number, numType);
        Assert.Equal("123", numVal.Text);

        var (strVal, strType) = CoerceValueHelper.Coerce("Hello", s => new CellValue("IDX"));
        Assert.Equal(CellValues.SharedString, strType);
        Assert.Equal("IDX", strVal.Text);
    }

    [Fact]
    public void Coerce_LargeIntegers_PreservePrecision()
    {
        var (longValue, longType) = CoerceValueHelper.Coerce(long.MaxValue, s => new CellValue(s));
        Assert.Equal(CellValues.Number, longType);
        Assert.Equal(long.MaxValue.ToString(CultureInfo.InvariantCulture), longValue.Text);

        var (ulongValue, ulongType) = CoerceValueHelper.Coerce(ulong.MaxValue, s => new CellValue(s));
        Assert.Equal(CellValues.Number, ulongType);
        Assert.Equal(ulong.MaxValue.ToString(CultureInfo.InvariantCulture), ulongValue.Text);
    }

    [Fact]
    public void Coerce_StringLengthBoundary()
    {
        string? captured = null;
        string withinLimit = new('a', 32_767);

        var (cellValue, type) = CoerceValueHelper.Coerce(withinLimit, s =>
        {
            captured = s;
            return new CellValue("IDX");
        });

        Assert.Equal(CellValues.SharedString, type);
        Assert.Equal("IDX", cellValue.Text);
        Assert.Equal(withinLimit, captured);

        string beyondLimit = new('b', 32_768);
        var exception = Assert.Throws<ArgumentException>(() => CoerceValueHelper.Coerce(beyondLimit, s => new CellValue(s)));
        Assert.Equal("value", exception.ParamName);
    }

    [Fact]
    public void Coerce_NullAndDbNull_ReturnEmptyString()
    {
        var (nullValue, nullType) = CoerceValueHelper.Coerce(null, _ => throw new InvalidOperationException("Should not use shared strings"));
        Assert.Equal(CellValues.String, nullType);
        Assert.Equal(string.Empty, nullValue.Text);

        var (dbNullValue, dbNullType) = CoerceValueHelper.Coerce(DBNull.Value, _ => throw new InvalidOperationException("Should not use shared strings"));
        Assert.Equal(CellValues.String, dbNullType);
        Assert.Equal(string.Empty, dbNullValue.Text);
    }

    [Fact]
    public void Coerce_ReferenceTypes_UseSharedStrings()
    {
        (CellValue value, CellValues type, string captured) Execute(object input)
        {
            string? seen = null;
            var result = CoerceValueHelper.Coerce(input, s =>
            {
                seen = s;
                return new CellValue("IDX");
            });

            return (result.cellValue, result.type, seen!);
        }

        var guid = Guid.NewGuid();
        var (guidValue, guidType, guidCaptured) = Execute(guid);
        Assert.Equal(CellValues.SharedString, guidType);
        Assert.Equal("IDX", guidValue.Text);
        Assert.Equal(guid.ToString(), guidCaptured);

        var (enumValue, enumType, enumCaptured) = Execute(DayOfWeek.Monday);
        Assert.Equal(CellValues.SharedString, enumType);
        Assert.Equal("IDX", enumValue.Text);
        Assert.Equal(DayOfWeek.Monday.ToString(), enumCaptured);

        var (charValue, charType, charCaptured) = Execute('X');
        Assert.Equal(CellValues.SharedString, charType);
        Assert.Equal("IDX", charValue.Text);
        Assert.Equal("X", charCaptured);

        var uri = new Uri("https://officeimo.net");
        var (uriValue, uriType, uriCaptured) = Execute(uri);
        Assert.Equal(CellValues.SharedString, uriType);
        Assert.Equal("IDX", uriValue.Text);
        Assert.Equal(uri.ToString(), uriCaptured);
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void Coerce_DateOnlyAndTimeOnly_AreNumeric()
    {
        var dateOnly = new DateOnly(2024, 3, 1);
        var expectedDate = dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture);
        var (dateValue, dateType) = CoerceValueHelper.Coerce(dateOnly, s => new CellValue(s));
        Assert.Equal(CellValues.Number, dateType);
        Assert.Equal(expectedDate, dateValue.Text);

        var timeOnly = new TimeOnly(13, 45, 30);
        var expectedTime = timeOnly.ToTimeSpan().TotalDays.ToString(CultureInfo.InvariantCulture);
        var (timeValue, timeType) = CoerceValueHelper.Coerce(timeOnly, s => new CellValue(s));
        Assert.Equal(CellValues.Number, timeType);
        Assert.Equal(expectedTime, timeValue.Text);
    }
#endif
}
