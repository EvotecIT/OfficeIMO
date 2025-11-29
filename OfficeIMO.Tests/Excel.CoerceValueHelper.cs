using System;
using System.Collections.Generic;
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

    [Fact]
    public void Coerce_DateTimeOffset_UsesLocalStrategyByDefault()
    {
        var offset = new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(5));
        var expected = offset.LocalDateTime.ToOADate().ToString(CultureInfo.InvariantCulture);

        var (value, type) = CoerceValueHelper.Coerce(offset, s => new CellValue(s));

        Assert.Equal(CellValues.Number, type);
        Assert.Equal(expected, value.Text);
    }

    [Fact]
    public void Coerce_DateTimeOffset_AllowsCustomStrategy()
    {
        var offset = new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(-3));
        var expected = offset.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture);

        var (value, type) = CoerceValueHelper.Coerce(offset, s => new CellValue(s), dto => dto.UtcDateTime);

        Assert.Equal(CellValues.Number, type);
        Assert.Equal(expected, value.Text);
    }

    public static IEnumerable<object[]> DateTimeOffsetSerialExpectations()
    {
        var customEastern = CreateCustomEasternTimeZone();

        yield return new object[]
        {
            new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.Zero),
            (Func<DateTimeOffset, DateTime>)(dto => dto.UtcDateTime),
            "45293.12783564815",
            "UTC timestamps should serialize to the exact documented OADate representation."
        };

        yield return new object[]
        {
            new DateTimeOffset(2024, 3, 10, 6, 30, 0, TimeSpan.Zero),
            CreateZoneStrategy(customEastern),
            "45361.0625",
            "Local conversion before the US DST jump preserves 01:30 standard time as serial 45361.0625."
        };

        yield return new object[]
        {
            new DateTimeOffset(2024, 3, 10, 7, 30, 0, TimeSpan.Zero),
            CreateZoneStrategy(customEastern),
            "45361.145833333336",
            "Local conversion after the DST gap advances to 03:30 and must emit serial 45361.145833333336."
        };

        yield return new object[]
        {
            new DateTimeOffset(2024, 11, 3, 5, 30, 0, TimeSpan.Zero),
            CreateZoneStrategy(customEastern),
            "45599.0625",
            "During the repeated hour at DST fall-back, 01:30 should serialize to 45599.0625."
        };

        yield return new object[]
        {
            new DateTimeOffset(2024, 5, 1, 10, 0, 0, TimeSpan.FromHours(2)),
            (Func<DateTimeOffset, DateTime>)(dto => dto.DateTime),
            "45413.416666666664",
            "Custom strategies that ignore offsets must continue to produce the documented wall-clock serial."
        };
    }

    [Theory]
    [MemberData(nameof(DateTimeOffsetSerialExpectations))]
    public void Coerce_DateTimeOffset_DocumentsSerialValues(
        DateTimeOffset input,
        Func<DateTimeOffset, DateTime> strategy,
        string expectedSerial,
        string reason)
    {
        var (value, type) = CoerceValueHelper.Coerce(input, s => new CellValue(s), strategy);

        Assert.Equal(CellValues.Number, type);
        Assert.Equal(expectedSerial, value.Text);
        Assert.True(expectedSerial == value.Text, reason);
    }

    private static Func<DateTimeOffset, DateTime> CreateZoneStrategy(TimeZoneInfo zone) =>
        dto => TimeZoneInfo.ConvertTime(dto, zone).DateTime;

    private static TimeZoneInfo CreateCustomEasternTimeZone()
    {
        var daylightStart = TimeZoneInfo.TransitionTime.CreateFloatingDateRule(new DateTime(1, 1, 1, 2, 0, 0), 3, 2, DayOfWeek.Sunday);
        var daylightEnd = TimeZoneInfo.TransitionTime.CreateFloatingDateRule(new DateTime(1, 1, 1, 2, 0, 0), 11, 1, DayOfWeek.Sunday);
        var adjustment = TimeZoneInfo.AdjustmentRule.CreateAdjustmentRule(
            new DateTime(2007, 1, 1),
            DateTime.MaxValue.Date,
            TimeSpan.FromHours(1),
            daylightStart,
            daylightEnd);

        return TimeZoneInfo.CreateCustomTimeZone(
            "OfficeIMO Eastern",
            TimeSpan.FromHours(-5),
            "OfficeIMO Eastern",
            "OfficeIMO Eastern",
            "OfficeIMO Eastern Daylight",
            new[] { adjustment });
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
