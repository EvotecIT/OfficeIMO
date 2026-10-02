using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Project;
using Xunit;

namespace OfficeIMO.Project.Tests;

public class ProjectTimePrecisionTests {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/project";

    [Theory]
    [InlineData("PT1M", 7)]
    [InlineData("PT61M", 9)]
    [InlineData("PT1S", 5)]
    [InlineData("PT0.0000001S", 11)]
    [InlineData("PT1M", 8)]
    [InlineData("PT1M", 39)]
    public void ImportedDurationSurvivesUnchangedAndEditedXmlSave(string text, int format) {
        byte[] source = Encoding.UTF8.GetBytes($"<Project xmlns='{Ns}'><MinutesPerDay>480</MinutesPerDay><Tasks><Task><UID>1</UID><Name>Original</Name><Duration>{text}</Duration><DurationFormat>{format}</DurationFormat></Task></Tasks></Project>");
        using var document = ProjectDocument.Load(new MemoryStream(source));
        Assert.False(document.IsModified); document.AssessSave().ThrowIfErrors();
        using var unchanged = new MemoryStream(); document.Save(unchanged); Assert.Equal(source, unchanged.ToArray());
        document.Tasks[0].Name = "Edited";
        using var edited = new MemoryStream(); document.Save(edited);
        Assert.Equal(XmlConvert.ToTimeSpan(text), XmlConvert.ToTimeSpan(XDocument.Parse(Encoding.UTF8.GetString(edited.ToArray())).Descendants(Ns + "Duration").Single().Value));
        using var reopened = ProjectDocument.Load(new MemoryStream(edited.ToArray())); reopened.AssessSave().ThrowIfErrors();
        Assert.Equal("Edited", reopened.Tasks[0].Name);
        Assert.Equal(document.Tasks[0].Duration, reopened.Tasks[0].Duration);
    }

    [Fact]
    public void ImportedWorkAcrossOwnersSurvivesAnEditedSave() {
        const string xml = "<Project xmlns='http://schemas.microsoft.com/project'><Tasks><Task><UID>1</UID><Work>PT1S</Work><Baseline><Number>0</Number><Work>PT0.0000001S</Work></Baseline></Task></Tasks><Resources><Resource><UID>1</UID><Work>PT1S</Work><Baseline><Number>0</Number><Work>PT1S</Work></Baseline></Resource></Resources><Assignments><Assignment><UID>1</UID><TaskUID>1</TaskUID><ResourceUID>1</ResourceUID><OvertimeWork>PT1S</OvertimeWork><Baseline><Number>0</Number><Work>PT1S</Work></Baseline></Assignment></Assignments></Project>";
        using var document = ProjectDocument.Load(new MemoryStream(Encoding.UTF8.GetBytes(xml)));
        document.Title = "Edited"; document.AssessSave().ThrowIfErrors();
        using var output = new MemoryStream(); document.Save(output);
        var root = XDocument.Parse(Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(new[] { TimeSpan.FromSeconds(1), TimeSpan.FromTicks(1), TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1) },
            root.Descendants(Ns + "Work").Select(e => XmlConvert.ToTimeSpan(e.Value)));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(root.Descendants(Ns + "OvertimeWork").Single().Value));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray())); reopened.AssessSave().ThrowIfErrors();
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void DayDisplayDurationSurvivesNativeConversionAndFieldEdit(ProjectFileFormat format) {
        using var authored = ProjectDocument.Create();
        authored.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
        var task = authored.Tasks.Add("Minute"); task.Duration = ProjectDuration.WorkingMinutes(1); task.RemainingDuration = task.Duration;
        string xml = authored.ToXml().Replace("<DurationFormat>3</DurationFormat>", "<DurationFormat>7</DurationFormat>");
        using var imported = ProjectDocument.Parse(xml);
        var options = new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };
        imported.AssessSave(options).ThrowIfErrors();
        using var first = new MemoryStream(); imported.Save(first, options);
        using var native = ProjectDocument.Load(new MemoryStream(first.ToArray()));
        native.Tasks[0].Name = "Edited minute";
        using var second = new MemoryStream(); native.Save(second, options);
        using var reopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Equal(ProjectDurationUnit.Day, reopened.Tasks[0].Duration!.Value.Unit);
        reopened.Settings.CurrencyCode = "USD";
        string exported = reopened.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Equal(TimeSpan.FromMinutes(1), XmlConvert.ToTimeSpan(XDocument.Parse(exported).Descendants(Ns + "Duration").Single().Value));
    }

    [Theory]
    [InlineData("PT1M", 7)]
    [InlineData("PT1S", 3)]
    [InlineData("PT1M", 8)]
    [InlineData("PT1M", 39)]
    public void MpxCustomDurationOutputReopensWithItsExactTimeAndFormat(string text, int format) {
        using var document = ProjectDocument.Create();
        var custom = document.Tasks.Add("Custom duration").CustomFields.Add();
        custom.FieldId = "188743783"; custom.Value = text; custom.DurationFormat = format;
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4, LossPolicy = OfficeConversionLossPolicy.Allow };
        document.AssessSave(options).ThrowIfErrors(); using var output = new MemoryStream(); document.Save(output, options);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var value = Assert.Single(reopened.Tasks[0].CustomFields);
        Assert.Equal(XmlConvert.ToTimeSpan(text), XmlConvert.ToTimeSpan(value.Value!)); Assert.Equal(format, value.DurationFormat);
    }

    [Fact]
    public void ImportedDurationTracksWorkingUnitSettingChangesAndEstimateMarkers() {
        using var document = ProjectDocument.Parse("<Project xmlns='http://schemas.microsoft.com/project'><MinutesPerDay>480</MinutesPerDay><Tasks><Task><UID>1</UID><Duration>PT1M</Duration><DurationFormat>7</DurationFormat></Task></Tasks></Project>");
        document.Settings.MinutesPerDay = 600; document.Tasks[0].Duration = document.Tasks[0].Duration!.Value.Estimated();
        var xml = XDocument.Parse(document.ToXml());
        Assert.Equal(TimeSpan.FromSeconds(75), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Duration").Single().Value));
        Assert.Equal("39", xml.Descendants(Ns + "DurationFormat").Single().Value);
    }

    [Fact]
    public void ImportedDayDisplayLagSurvivesXmlAndNativeSave() {
        using var document = ProjectDocument.Parse("<Project xmlns='http://schemas.microsoft.com/project'><MinutesPerDay>480</MinutesPerDay><Tasks><Task><UID>1</UID><Name>First</Name></Task><Task><UID>2</UID><Name>Second</Name><PredecessorLink><PredecessorUID>1</PredecessorUID><Type>1</Type><LinkLag>10</LinkLag><LagFormat>7</LagFormat></PredecessorLink></Task></Tasks></Project>");
        document.Title = "Edited";
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        Assert.Equal("10", XDocument.Parse(document.ToXml()).Descendants(Ns + "LinkLag").Single().Value);
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14, LossPolicy = OfficeConversionLossPolicy.Allow };
        document.AssessSave(options).ThrowIfErrors();
        using var output = new MemoryStream(); document.Save(output, options);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(document.Dependencies[0].Lag, reopened.Dependencies[0].Lag);
    }

    [Fact]
    public void AppliedSchedulePreservesImportedSecondWorkAndDayDisplayDuration() {
        using var document = ProjectDocument.Parse("<Project xmlns='http://schemas.microsoft.com/project'><StartDate>2026-10-05T08:00:00</StartDate><MinutesPerDay>480</MinutesPerDay><Tasks><Task><UID>1</UID><Name>Second</Name><Duration>PT1S</Duration><DurationFormat>7</DurationFormat><Work>PT1S</Work></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 60;
        var assignment = document.Assignments.Add(document.Tasks[0], resource, ProjectUnits.Percent(100)); assignment.Work = document.Tasks[0].Work;
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        document.AssessSave(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }).ThrowIfErrors();
        var xml = XDocument.Parse(document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Duration").Single().Value));
        Assert.All(xml.Descendants(Ns + "Work"), value => Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(value.Value)));
        Assert.All(xml.Descendants(Ns + "RemainingWork"), value => Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(value.Value)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TickExactWorkSurvivesTaskSummaryResourceAndBaselineAggregation(bool oneTask) {
        string tasks = string.Concat(Enumerable.Range(1, oneTask ? 1 : 3).Select(uid => $"<Task><UID>{uid}</UID><Name>Second {uid}</Name><OutlineLevel>2</OutlineLevel><Duration>PT1S</Duration><DurationFormat>3</DurationFormat><Work>PT1S</Work></Task>"));
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>10</UID><Name>Summary</Name><OutlineLevel>1</OutlineLevel><Summary>1</Summary></Task>{tasks}</Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var shared = document.Resources.AddWork("Shared"); shared.StandardRate = 60;
        for (int index = 0; index < 3; index++) {
            var task = document.Tasks.GetByUid(oneTask ? 1 : index + 1);
            var resource = oneTask && index > 0 ? document.Resources.AddWork("Engineer " + index) : shared; resource.StandardRate = 60;
            document.Assignments.Add(task, resource, ProjectUnits.Percent(100)).Work = task.Work;
        }
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        using var output = new MemoryStream(); document.Save(output, new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var xml = XDocument.Parse(Encoding.UTF8.GetString(output.ToArray()));
        var summary = xml.Descendants(Ns + "Task").Single(t => t.Element(Ns + "UID")!.Value == "10");
        Assert.Equal(TimeSpan.FromSeconds(3), XmlConvert.ToTimeSpan(summary.Element(Ns + "Work")!.Value));
        Assert.Equal(TimeSpan.FromSeconds(3), XmlConvert.ToTimeSpan(summary.Element(Ns + "Baseline")!.Element(Ns + "Work")!.Value));
        var resourceXml = xml.Descendants(Ns + "Resource").Single(r => r.Element(Ns + "UID")!.Value == shared.Uid.ToString());
        Assert.Equal(TimeSpan.FromSeconds(oneTask ? 1 : 3), XmlConvert.ToTimeSpan(resourceXml.Element(Ns + "Work")!.Value));
        Assert.Equal(TimeSpan.FromSeconds(oneTask ? 1 : 3), XmlConvert.ToTimeSpan(resourceXml.Element(Ns + "Baseline")!.Element(Ns + "Work")!.Value));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray())); reopened.AssessSave().ThrowIfErrors();
    }

    [Fact]
    public void ExactSecondTimesSurviveMappedDataAndScaledWorkEquations() {
        using var source = ProjectDocument.Parse($"<Project xmlns='{Ns}'><Tasks><Task><UID>1</UID><Name>Second</Name><Duration>PT1S</Duration><DurationFormat>7</DurationFormat><Work>PT1S</Work></Task></Tasks></Project>");
        using var imported = ProjectDocument.ImportTables(source.ExportTables(allowLossyProjection: true).Tables).Document;
        var xml = XDocument.Parse(imported.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Duration").Single().Value));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Work").Single().Value));
        var scaled = ProjectWorkEquation.Work(source.Tasks[0].Work!.Value.Minutes, ProjectUnits.Percent(300));
        Assert.Equal(3m, ProjectWorkEquation.Units(scaled, source.Tasks[0].Work!.Value.Minutes).Value);
        imported.Tasks[0].Work = scaled;
        xml = XDocument.Parse(imported.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(TimeSpan.FromSeconds(3), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Work").Single().Value));
    }

    [Fact]
    public void UnassignedTaskProgressCombinesExactActualAndRemainingWork() {
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>1</UID><Name>Progress</Name><Duration>PT2S</Duration><DurationFormat>3</DurationFormat><ActualDuration>PT1S</ActualDuration><RemainingDuration>PT1S</RemainingDuration><ActualStart>2026-10-05T08:00:00</ActualStart><ActualWork>PT1S</ActualWork><RemainingWork>PT1S</RemainingWork></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        var xml = XDocument.Parse(document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(TimeSpan.FromSeconds(2), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "Work").Single().Value));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "ActualWork").Single().Value));
        Assert.Equal(TimeSpan.FromSeconds(1), XmlConvert.ToTimeSpan(xml.Descendants(Ns + "RemainingWork").Single().Value));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AssignedSecondProgressInfersExactRemainingWorkAndOvertime(bool overtime) {
        int work = overtime ? 4 : 2, actual = overtime ? 2 : 1;
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>1</UID><Name>Progress</Name><Duration>PT2S</Duration><DurationFormat>3</DurationFormat><ActualDuration>PT1S</ActualDuration><RemainingDuration>PT1S</RemainingDuration><ActualStart>2026-10-05T08:00:00</ActualStart><Work>PT{work}S</Work><ActualWork>PT{actual}S</ActualWork></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 60; resource.OvertimeRate = 60;
        var assignment = document.Assignments.Add(document.Tasks[0], resource, ProjectUnits.Percent(100));
        assignment.Work = document.Tasks[0].Work; assignment.ActualWork = document.Tasks[0].ActualWork; assignment.ActualStart = document.Tasks[0].ActualStart;
        if (overtime) {
            using var times = ProjectDocument.Parse($"<Project xmlns='{Ns}'><Tasks><Task><UID>1</UID><Work>PT2S</Work><ActualWork>PT1S</ActualWork></Task></Tasks></Project>");
            assignment.OvertimeWork = times.Tasks[0].Work; assignment.ActualOvertimeWork = times.Tasks[0].ActualWork;
        }
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        var xml = XDocument.Parse(document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var values = xml.Descendants(Ns + "Assignment").Single();
        Assert.Equal(TimeSpan.FromSeconds(work), XmlConvert.ToTimeSpan(values.Element(Ns + "Work")!.Value));
        Assert.Equal(TimeSpan.FromSeconds(actual), XmlConvert.ToTimeSpan(values.Element(Ns + "ActualWork")!.Value));
        Assert.Equal(TimeSpan.FromSeconds(work - actual), XmlConvert.ToTimeSpan(values.Element(Ns + "RemainingWork")!.Value));
        if (overtime) Assert.Equal(TimeSpan.FromSeconds(2), XmlConvert.ToTimeSpan(values.Element(Ns + "OvertimeWork")!.Value));
    }

    [Theory]
    [InlineData("PT2S", "PT1S", "PT1S", 1, 1)]
    [InlineData("PT1H", "PT20M", "PT40M", 1200, 2400)]
    public void SummaryProgressScalesExactChildDuration(string total, string actual, string remaining, int actualSeconds, int remainingSeconds) {
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>10</UID><Name>Summary</Name><OutlineLevel>1</OutlineLevel><Summary>1</Summary></Task><Task><UID>1</UID><Name>Progress</Name><OutlineLevel>2</OutlineLevel><Duration>{total}</Duration><DurationFormat>3</DurationFormat><ActualDuration>{actual}</ActualDuration><RemainingDuration>{remaining}</RemainingDuration><ActualStart>2026-10-05T08:00:00</ActualStart></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        var xml = XDocument.Parse(document.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var summary = xml.Descendants(Ns + "Task").Single(t => t.Element(Ns + "UID")!.Value == "10");
        Assert.Equal(TimeSpan.FromSeconds(actualSeconds), XmlConvert.ToTimeSpan(summary.Element(Ns + "ActualDuration")!.Value));
        Assert.Equal(TimeSpan.FromSeconds(remainingSeconds), XmlConvert.ToTimeSpan(summary.Element(Ns + "RemainingDuration")!.Value));
    }

    [Fact]
    public void ActualCurveProportionsPreserveExactWorkThroughBaselineCapture() {
        var start = new DateTime(2026, 10, 5, 8, 0, 0);
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>1</UID><Name>Recorded curve</Name><Duration>PT3S</Duration><DurationFormat>3</DurationFormat><ActualDuration>PT3S</ActualDuration><RemainingDuration>PT0S</RemainingDuration><ActualStart>2026-10-05T08:00:00</ActualStart><Work>PT1M</Work><ActualWork>PT1M</ActualWork></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks[0];
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = resource.OvertimeRate = 60;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Percent(100));
        assignment.Work = assignment.ActualWork = task.Work;
        using var second = ProjectDocument.Parse($"<Project xmlns='{Ns}'><Tasks><Task><UID>1</UID><Work>PT1S</Work></Task></Tasks></Project>");
        assignment.OvertimeWork = assignment.ActualOvertimeWork = second.Tasks[0].Work;
        foreach (var spec in new[] { (Type: 2, Seconds: 3, Value: "PT1M"), (Type: 3, Seconds: 1, Value: "PT1S") }) {
            var curve = assignment.TimephasedData.Add(); curve.Type = spec.Type; curve.Uid = assignment.Uid;
            curve.Start = start; curve.Finish = start.AddSeconds(spec.Seconds); curve.Value = spec.Value;
        }
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        schedule.Report.ThrowIfErrors();
        Assert.Equal(new[] { 20m / 60m, 40m / 60m }, schedule.Assignments.Single().Intervals.Select(v => v.Work.Minutes).ToArray());
        document.CaptureBaseline(schedule);
        var baseline = assignment.TimephasedData.Where(v => v.Type == 4).OrderBy(v => v.Start).ToArray();
        Assert.Equal(new[] { TimeSpan.FromSeconds(20), TimeSpan.FromSeconds(40) }, baseline.Select(v => XmlConvert.ToTimeSpan(v.Value!)).ToArray());
        document.AssessSave(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }).ThrowIfErrors();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactHalfPercentCompletionRoundsAwayFromZero(bool assigned) {
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><StartDate>2026-10-05T08:00:00</StartDate><Tasks><Task><UID>1</UID><Name>Progress</Name><Duration>PT4S</Duration><DurationFormat>3</DurationFormat><ActualDuration>PT1.98S</ActualDuration><RemainingDuration>PT2.02S</RemainingDuration><ActualStart>2026-10-05T08:00:00</ActualStart><Work>PT4S</Work><ActualWork>PT1.98S</ActualWork><RemainingWork>PT2.02S</RemainingWork></Task></Tasks></Project>");
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks[0];
        if (assigned) {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"), ProjectUnits.Percent(100));
            assignment.Work = task.Work; assignment.ActualWork = task.ActualWork; assignment.RemainingWork = task.RemainingWork; assignment.ActualStart = task.ActualStart;
        }
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(50, task.PercentComplete); Assert.Equal(50, task.PercentWorkComplete);
        if (assigned) Assert.Equal(50, document.Assignments.Single().PercentWorkComplete);
    }

    [Fact]
    public void StoredAnalysisBalancesAndAggregatesExactSecondQuantities() {
        string tasks = string.Concat(Enumerable.Range(1, 3).Select(uid => $"<Task><UID>{uid}</UID><Name>Progress {uid}</Name><Work>PT2S</Work><ActualWork>PT1S</ActualWork><RemainingWork>PT1S</RemainingWork></Task>"));
        using var document = ProjectDocument.Parse($"<Project xmlns='{Ns}'><Tasks>{tasks}</Tasks></Project>");
        var resource = document.Resources.AddWork("Shared");
        foreach (var task in document.Tasks) {
            var assignment = document.Assignments.Add(task, resource);
            assignment.Work = task.Work; assignment.ActualWork = task.ActualWork; assignment.RemainingWork = task.RemainingWork;
        }
        var result = document.AnalyzeAssignments();
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "PROJECT_WORK_BALANCE");
        var totals = Assert.Single(result.Resources);
        Assert.Equal(.1m, totals.WorkMinutes); Assert.Equal(.05m, totals.ActualWorkMinutes); Assert.Equal(.05m, totals.RemainingWorkMinutes);
    }
}
