using OfficeIMO.Core.Internal;
using System.Collections.Generic;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectFormatRegressionTests {
    [Theory]
    [InlineData(ProjectFileFormat.Xml)]
    [InlineData(ProjectFileFormat.Mpx4)]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpt8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpt9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpt12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    [InlineData(ProjectFileFormat.Mpt14)]
    public void FullDayCalendarsRetainTwentyFourHoursAcrossEveryFormat(ProjectFileFormat format) {
        using var document = ProjectDocument.Create();
        document.Settings.StartDate = new DateTime(2026, 10, 5);
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek))) document.Calendar.SetWorkingDay(day, ProjectWorkingTime.Hours(0, 0));
        var task = document.Tasks.Add("Continuous operation"); task.Duration = ProjectDuration.WorkingHours(24);
        var options = new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, options);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.DoesNotContain(reopened.Validate().Diagnostics, d => d.Code == "PROJECT_WORKING_INTERVAL");
        DateTime monday = new DateTime(2026, 10, 5);
        var interval = Assert.Single(reopened.Calendar!.GetWorkingIntervals(monday));
        Assert.Equal(monday, interval.Start); Assert.Equal(monday.AddDays(1), interval.Finish);
        Assert.Equal(monday.AddDays(2), reopened.Calendar.AddWorkingMinutes(monday, 2880));
        reopened.Tasks.GetByUid(task.Uid).Name = "Edited continuous operation";
        using var edited = new MemoryStream(); reopened.Save(edited, options);
        using var again = ProjectDocument.Load(new MemoryStream(edited.ToArray()));
        Assert.Equal(monday.AddDays(1), Assert.Single(again.Calendar!.GetWorkingIntervals(monday)).Finish);
    }

    [Fact]
    public void LegacyDuplicateTrailingMetadataDoesNotCreateEntitiesAndRemainsEditable() {
        using var document = ProjectNativeAuthoringTests.Create();
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp9, LossPolicy = OfficeConversionLossPolicy.Allow };
        using var output = new MemoryStream(); document.Save(output, options);
        Assert.True(OfficeCompoundFileReader.TryRead(output.ToArray(), out OfficeCompoundFile? compound, out var error), error);
        var replacements = new Dictionary<string, byte[]>();
        foreach (string table in new[] { "Task", "Assn" }) {
            string path = "   19/TBknd" + table + "/FixedMeta";
            byte[] original = compound!.Streams[path]; int count = BitConverter.ToInt32(original, 8);
            int width = (original.Length - 16) / count;
            replacements.Add(path, original.Concat(original.Skip(original.Length - width)).ToArray());
        }
        byte[] source = OfficeCompoundFileWriter.Rewrite(compound!, replacements);
        using var reopened = ProjectDocument.Load(new MemoryStream(source));
        Assert.Single(reopened.Assignments);
        Assert.Equal("Design café / Łódź / 日本語", reopened.Tasks.GetByUid(2).Name);
        using var unchanged = new MemoryStream(); reopened.Save(unchanged);
        Assert.Equal(source, unchanged.ToArray());
        reopened.Tasks.GetByUid(2).Name = "Expanded task name";
        var added = reopened.Tasks.Add("New task");
        using var edited = new MemoryStream(); reopened.Save(edited, options);
        using var again = ProjectDocument.Load(new MemoryStream(edited.ToArray()));
        Assert.Equal("Expanded task name", again.Tasks.GetByUid(2).Name);
        Assert.Equal("New task", again.Tasks.GetByUid(added.Uid).Name);
        Assert.Single(again.Assignments);
    }
}
