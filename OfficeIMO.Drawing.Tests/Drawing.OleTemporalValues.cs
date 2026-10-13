using System;
using System.IO;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingOleTemporalValueTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidTemporalValuesAreMalformedPropertyData(bool oleDate) {
        OfficeOleProperty property = TemporalProperty(oleDate, invalid: true);
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetReader.ReadSections(PropertySet(property)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValidTemporalValuesRetainTheirUtcMeaning(bool oleDate) {
        DateTime expected = oleDate
            ? DateTime.FromOADate(45000).ToUniversalTime()
            : DateTime.FromFileTimeUtc(132537600000000000);
        var sections = OfficeOlePropertySetReader.ReadSections(PropertySet(TemporalProperty(oleDate, invalid: false)));
        Assert.Equal(expected, Assert.IsType<DateTime>(Assert.Single(sections).Properties[12].Value));
    }

    private static OfficeOleProperty TemporalProperty(bool oleDate, bool invalid) {
        OfficeOleProperty property = oleDate
            ? OfficeOleProperty.Double(12, invalid ? double.NaN : 45000)
            : OfficeOleProperty.Int64(12, invalid ? -1 : 132537600000000000);
        property.ValueBytes[0] = oleDate ? (byte)7 : (byte)0x40;
        return property;
    }

    private static byte[] PropertySet(OfficeOleProperty property) =>
        OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId,
            OfficeOlePropertySetWriter.CreateSection(new[] { property })));
}
