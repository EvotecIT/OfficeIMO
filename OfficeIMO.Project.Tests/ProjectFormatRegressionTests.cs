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
    public void RetainedOutputAssessmentHonorsTheExactByteLimit(ProjectFileFormat format) {
        using var authored = ProjectDocument.Create(); authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
        authored.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); authored.Tasks.Add("Retained document");
        using var source = new MemoryStream();
        authored.Save(source, new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow });
        byte[] bytes = source.ToArray();
        using var document = ProjectDocument.Load(new MemoryStream(bytes));
        var exact = new ProjectSaveOptions { Format = format, MaxOutputBytes = bytes.Length, LossPolicy = OfficeConversionLossPolicy.Allow };
        document.AssessSave(exact).ThrowIfErrors();
        using var copy = new MemoryStream(); document.Save(copy, exact); Assert.Equal(bytes, copy.ToArray());
        var tooSmall = new ProjectSaveOptions { Format = format, MaxOutputBytes = bytes.Length - 1, LossPolicy = OfficeConversionLossPolicy.Allow };
        Assert.True(document.AssessSave(tooSmall).HasErrors);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => document.Save(destination, tooSmall));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        Assert.False(document.IsModified);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedXmlAssessmentHonorsTheExactByteLimit(bool editLoadedDocument) {
        using var authored = ProjectDocument.Create(); authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
        authored.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); authored.Tasks.Add("Generated document");
        ProjectDocument document;
        if (editLoadedDocument) {
            byte[] source = ProjectXmlCodec.Write(authored, new ProjectSaveOptions { Format = ProjectFileFormat.Xml }, default);
            document = ProjectDocument.Load(new MemoryStream(source));
            document.Tasks[0].Name = "Edited generated document";
        } else {
            document = authored;
        }
        using (document == authored ? null : document) {
            var generatedOptions = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, PreserveUnchangedBytes = false, LossPolicy = OfficeConversionLossPolicy.Allow };
            byte[] generated = ProjectXmlCodec.Write(document, generatedOptions, default);
            var exact = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, PreserveUnchangedBytes = false, MaxOutputBytes = generated.Length, LossPolicy = OfficeConversionLossPolicy.Allow };
            document.AssessSave(exact).ThrowIfErrors();
            var tooSmall = new ProjectSaveOptions { Format = ProjectFileFormat.Xml, PreserveUnchangedBytes = false, MaxOutputBytes = generated.Length - 1, LossPolicy = OfficeConversionLossPolicy.Allow };
            var assessment = document.AssessSave(tooSmall);
            Assert.Contains(assessment.Diagnostics, d => d.Code == "PROJECT_OUTPUT_LIMIT" && d.Severity == ProjectDiagnosticSeverity.Error);
            using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
            Assert.Throws<InvalidDataException>(() => document.Save(destination, tooSmall));
            Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        }
    }

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
