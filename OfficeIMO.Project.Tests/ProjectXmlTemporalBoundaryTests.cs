namespace OfficeIMO.Project.Tests;

public sealed class ProjectXmlTemporalBoundaryTests {
    [Theory]
    [InlineData("2026-10-05T08:00:00Z")]
    [InlineData("2026-10-05T08:00:00+02:00")]
    [InlineData("2026-10-05T08:00:00-07:00")]
    public void TimezoneQualifiedProjectDatesAreRejected(string value) {
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(XmlContracts.Wrap("<StartDate>" + value + "</StartDate>")));
        Assert.Contains("without a timezone suffix", error.Message);
    }

    [Theory]
    [InlineData("08:00:00Z")]
    [InlineData("08:00:00+02:00")]
    public void TimezoneQualifiedProjectClocksAreRejected(string value) {
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(XmlContracts.Wrap("<DefaultStartTime>" + value + "</DefaultStartTime>")));
        Assert.Contains("without a timezone suffix", error.Message);
    }

    [Fact]
    public void LocalFractionalProjectTimesRemainUnspecified() {
        using var document = ProjectDocument.Parse(XmlContracts.Wrap(
            "<StartDate>2026-10-05T08:00:00.1234567</StartDate><DefaultStartTime>08:00:00.1234567</DefaultStartTime>"));
        Assert.Equal(DateTimeKind.Unspecified, document.Settings.StartDate!.Value.Kind);
        Assert.Equal(new TimeSpan(0, 8, 0, 0, 123) + TimeSpan.FromTicks(4567), document.Settings.DefaultStartTime);
    }
}
