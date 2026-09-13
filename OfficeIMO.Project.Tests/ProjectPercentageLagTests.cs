using System.Xml.Linq;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectPercentageLagTests {
    [Theory]
    [InlineData(19)]
    [InlineData(20)]
    [InlineData(51)]
    [InlineData(52)]
    public void XmlRetainsPercentageLagFormatAndSchedulesItsOwnTimeBasis(int format) {
        using var source = ProjectDocument.Create(); source.Calendar = source.Calendars.AddStandardWorkingWeek();
        source.Settings.StartDate = new DateTime(2026, 10, 9, 8, 0, 0);
        var predecessor = source.Tasks.Add("Prepare"); predecessor.Duration = ProjectDuration.WorkingDays(1); predecessor.RemainingDuration = predecessor.Duration;
        var successor = source.Tasks.Add("Accept"); successor.Duration = ProjectDuration.WorkingMinutes(0); successor.RemainingDuration = successor.Duration;
        source.Dependencies.Add(predecessor, successor).LagPercent = 50;
        var xml = XDocument.Parse(source.ToXml()); var field = xml.Descendants().Single(e => e.Name.LocalName == "LagFormat"); field.Value = format.ToString();
        using var document = ProjectDocument.Parse(xml.ToString());
        Assert.Equal(format == 20 || format == 52, document.Dependencies.Single().LagPercentIsElapsed);
        Assert.Equal(format >= 51, document.Dependencies.Single().LagPercentIsEstimated);
        document.Dependencies.Single().LagPercent = 75;
        Assert.Equal(format.ToString(), XDocument.Parse(document.ToXml()).Descendants().Single(e => e.Name.LocalName == "LagFormat").Value);
        if (format >= 51) {
            var unsupported = document.CalculateSchedule(); Assert.True(unsupported.Report.HasErrors);
            Assert.Throws<InvalidDataException>(() => document.ApplySchedule(unsupported));
            document.Dependencies.Single().LagPercentIsEstimated = false;
        }
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        var finish = result.Tasks.Single(t => t.TaskUid == successor.Uid).Start;
        bool elapsed = format == 20 || format == 52;
        Assert.Equal(elapsed ? new DateTime(2026, 10, 9, 23, 0, 0) : new DateTime(2026, 10, 12, 15, 0, 0), finish);
        var view = document.CreateView(result);
        Assert.Equal("75" + (elapsed ? "e%" : "%"), Assert.Single(view.Links).LagText);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpt8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpt9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpt12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    [InlineData(ProjectFileFormat.Mpt14)]
    public void NativeAuthoringAndEditingRetainEveryPercentageLagFormat(ProjectFileFormat format) {
        foreach (int flags in new[] { 0, 1, 2, 3 }) {
            using var document = Create(); var link = document.Dependencies.Single();
            link.LagPercentIsElapsed = (flags & 1) != 0; link.LagPercentIsEstimated = (flags & 2) != 0;
            var options = new ProjectSaveOptions { Format = format };
            using var output = new MemoryStream(); document.Save(output, options); output.Position = 0;
            using var copy = ProjectDocument.Load(output); var copied = Assert.Single(copy.Dependencies);
            Assert.Equal(link.LagPercentIsElapsed, copied.LagPercentIsElapsed); Assert.Equal(link.LagPercentIsEstimated, copied.LagPercentIsEstimated);
            copied.LagPercent = -25;
            options.LossPolicy = OfficeConversionLossPolicy.Allow; // Native dependency edits retain opaque presentation references.
            using var changed = new MemoryStream(); copy.Save(changed, options); changed.Position = 0;
            using var reopened = ProjectDocument.Load(changed); var repeated = Assert.Single(reopened.Dependencies);
            Assert.Equal(-25m, repeated.LagPercent); Assert.Equal(link.LagPercentIsElapsed, repeated.LagPercentIsElapsed);
            Assert.Equal(link.LagPercentIsEstimated, repeated.LagPercentIsEstimated);
        }
    }

    [Fact]
    public void WorkingPercentageUsesSuccessorCalendarEvenWhenPredecessorIsElapsed() {
        using var document = Create(); document.Settings.StartDate = new DateTime(2026, 10, 9, 8, 0, 0);
        document.Tasks[0].Duration = new ProjectDuration(8, ProjectDurationUnit.Hour, true);
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(new DateTime(2026, 10, 12, 11, 0, 0), result.Tasks[1].Start);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MpxCannotSilentlyDiscardPercentageLagFlags(bool elapsed, bool estimated) {
        using var document = Create(); var link = document.Dependencies.Single(); link.LagPercentIsElapsed = elapsed; link.LagPercentIsEstimated = estimated;
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4 };
        Assert.Contains(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_MPX_PERCENTAGE_LAG_FORMAT");
        using var output = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => document.Save(output, options));
    }

    [Fact]
    public void LagFlagsInvalidateSchedulesAndAreClearedWithPercentage() {
        using var document = Create(); var result = document.CalculateSchedule(); var link = document.Dependencies.Single();
        link.LagPercentIsElapsed = true;
        Assert.Throws<InvalidOperationException>(() => document.CreateView(result));
        using var copied = document.Clone(); Assert.True(copied.Dependencies.Single().LagPercentIsElapsed);
        link.LagPercentIsEstimated = true; link.LagPercent = null;
        Assert.False(link.LagPercentIsElapsed); Assert.False(link.LagPercentIsEstimated);
        Assert.Throws<InvalidOperationException>(() => link.LagPercentIsElapsed = true);
        link.Lag = ProjectDuration.WorkingHours(1);
        Assert.Equal(ProjectDuration.WorkingHours(1), link.Lag);
    }

    private static ProjectDocument Create() {
        var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var first = document.Tasks.Add("Prepare"); first.Duration = ProjectDuration.WorkingDays(1); first.RemainingDuration = first.Duration;
        var second = document.Tasks.Add("Accept"); second.Duration = ProjectDuration.WorkingMinutes(0); second.RemainingDuration = second.Duration;
        first.IsManual = second.IsManual = false; first.IsActive = second.IsActive = true;
        first.IsNull = second.IsNull = false; first.IsMilestone = second.IsMilestone = false;
        first.IsCritical = second.IsCritical = false; first.EffortDriven = second.EffortDriven = false;
        document.Dependencies.Add(first, second).LagPercent = 50; return document;
    }
}
