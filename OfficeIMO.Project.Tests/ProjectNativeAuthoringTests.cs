namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeAuthoringTests {
    private static ProjectSaveOptions Native(ProjectFileFormat format = ProjectFileFormat.Mpp14, OfficeConversionLossPolicy policy = OfficeConversionLossPolicy.Block) =>
        new ProjectSaveOptions { Format = format, LossPolicy = policy };

    [Fact]
    public void NewNativeDocumentNeedsNoSourceAndRetainsItsMappedModel() {
        using var document = Create();
        Assert.Null(document.NativeInfo);
        using var output = new MemoryStream(); document.Save(output, Native());
        using var reload = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("MPP14", reload.NativeInfo!.Generation); Assert.False(reload.NativeInfo.IsTemplate);
        var task = reload.Tasks.GetByUid(2);
        Assert.Equal("Design café / Łódź / 日本語", task.Name); Assert.Equal(1, task.Parent!.Uid);
        Assert.Equal(ProjectDuration.WorkingDays(1), task.Duration); Assert.Equal(800m, task.Baselines.Single().Cost);
        Assert.Equal("Independent authoring", task.CustomFields.Single().Value);
        Assert.Equal("Work area", reload.CustomFields.Single().Alias);
        Assert.Equal(480m, reload.Assignments.Single().Work!.Value.Minutes);
        Assert.Equal("Engineer", reload.Assignments.Single().Resource!.Name);
        Assert.Equal("Maintenance", reload.Calendar!.Exceptions.Single().Name);
        Assert.Equal(new DateTime(2026, 10, 12), reload.Calendar.Exceptions.Single().FromDate);
        using var again = new MemoryStream(); document.Save(again, Native()); Assert.Equal(output.ToArray(), again.ToArray());
    }

    [Fact]
    public void XmlToNativeAndNativeToXmlHaveIndependentLossAssessment() {
        using var source = Create(); string xml = source.ToXml();
        using var parsed = ProjectDocument.Parse(xml);
        using var output = new MemoryStream(); parsed.Save(output, Native(policy: OfficeConversionLossPolicy.Allow));
        using var native = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var report = native.AssessSave(Native(ProjectFileFormat.Xml));
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_PRESENTATION_LOSS");
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_CURVE_LOSS");
        Assert.Throws<InvalidOperationException>(() => native.ToXml());
        using var converted = ProjectDocument.Parse(native.ToXml(Native(ProjectFileFormat.Xml, OfficeConversionLossPolicy.Allow)));
        Assert.Equal("Design café / Łódź / 日本語", converted.Tasks.GetByUid(2).Name);
        Assert.Equal(800m, converted.Tasks.GetByUid(2).Baselines.Single().Cost);
    }

    [Fact]
    public void TemplateCreationDoesNotAssociateTheSourceAsAnOutput() {
        using var source = Create(); using var bytes = new MemoryStream(); source.Save(bytes, Native(ProjectFileFormat.Mpt14));
        byte[] original = bytes.ToArray();
        using var template = ProjectDocument.Load(new MemoryStream(original)); Assert.True(template.NativeInfo!.IsTemplate);
        using var project = ProjectDocument.CreateFromTemplate(bytes);
        Assert.Throws<InvalidOperationException>(() => project.Save());
        project.Tasks.GetByUid(2).Name = "Template instance";
        using var output = new MemoryStream(); project.Save(output);
        using var reload = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.False(reload.NativeInfo!.IsTemplate); Assert.Equal("Template instance", reload.Tasks.GetByUid(2).Name);
        Assert.Equal(original, bytes.ToArray());
        Assert.Throws<NotSupportedException>(() => ProjectDocument.Create("Global.mpt"));
        Assert.Throws<ArgumentException>(() => source.AssessSave("mismatch.xml", Native()));
    }

    [Fact]
    public void AllowedOmissionsRemainVisibleOnSubsequentSavesOfTheCurrentModel() {
        using var document = Create(); document.Tasks.GetByUid(2).Notes = "Unqualified native notes";
        var report = document.AssessSave(Native());
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_FIELD_LOSS" && d.Location.EndsWith("/Notes", StringComparison.Ordinal));
        using var output = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(output, Native()));
        document.Save(output, Native(policy: OfficeConversionLossPolicy.Allow));
        Assert.True(document.AssessSave(Native()).HasLoss);
        using var second = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => document.Save(second, Native()));
        document.Tasks.GetByUid(2).Notes = null;
        Assert.DoesNotContain(document.AssessSave(Native()).Diagnostics, d => d.Code == "PROJECT_NATIVE_FIELD_LOSS");
        document.Save(second, Native());
    }

    [Fact]
    public void NativeCalendarOwnershipAndOutputBoundsAreExplicit() {
        using var document = Create(); var resource = document.Resources.GetByUid(1);
        resource.Calendar = document.Calendar;
        Assert.Contains(document.AssessSave(Native()).Diagnostics, d => d.Code == "PROJECT_NATIVE_RESOURCE_CALENDAR");
        using var target = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(target, Native())); Assert.Empty(target.ToArray());
        resource.Calendar = document.Calendars.Single(c => c.Name == "Engineer");
        Assert.Throws<InvalidDataException>(() => document.Save(target, new ProjectSaveOptions { Format = ProjectFileFormat.Mpp14, MaxOutputBytes = 128 }));
        Assert.Empty(target.ToArray());
    }

    [Fact]
    public void NativeAssessmentBoundsFindingsAndReportsInvalidCreationValues() {
        using var document = Create(); document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 1);
        Assert.Contains(document.AssessSave(Native()).Diagnostics, d => d.Severity == ProjectDiagnosticSeverity.Error);
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        document.Title = "Invalid\0title";
        Assert.Contains(document.AssessSave(Native()).Diagnostics, d => d.Code == "PROJECT_NATIVE_METADATA_VALUE");
        document.Title = "Valid";
        for (int i = 0; i < 1010; i++) document.Tasks.Add("Task " + i).Notes = "Unqualified native notes";
        var report = document.AssessSave(Native());
        Assert.True(report.HasErrors); Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_DIAGNOSTICS_TRUNCATED");
        Assert.InRange(report.Diagnostics.Count, 1001, 1005);
    }

    [Fact]
    public void DurationFormatsMustAgreeEvenWithoutMainDuration() {
        using var document = Create(); var task = document.Tasks.GetByUid(2);
        task.Duration = null; task.ActualDuration = ProjectDuration.WorkingHours(1); task.RemainingDuration = ProjectDuration.ElapsedHours(1);
        Assert.Contains(document.AssessSave(Native()).Diagnostics, d => d.Code == "PROJECT_NATIVE_DURATION_FORMAT");
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output, Native())); Assert.Empty(output.ToArray());
    }

    [Fact]
    public void DurationOnlyBaselinesSurviveNativeLoadAndClone() {
        using var document = Create(); var task = document.Tasks.GetByUid(2); task.Baselines.Remove(task.Baselines.Single());
        for (int i = 0; i <= 10; i++) { var baseline = task.Baselines.Add(); baseline.Number = i; baseline.Duration = ProjectDuration.WorkingDays(i + 1); }
        using var output = new MemoryStream(); document.Save(output, Native());
        using var loaded = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        using var clone = document.Clone();
        foreach (var result in new[] { loaded, clone }) {
            var baselines = result.Tasks.GetByUid(2).Baselines.OrderBy(b => b.Number).ToArray(); Assert.Equal(11, baselines.Length);
            for (int i = 0; i <= 10; i++) Assert.Equal(ProjectDuration.WorkingDays(i + 1), baselines[i].Duration);
        }
    }

    [Fact]
    public void CalculatedNativeCreationReportsUnmappedSlackBeforeAllowingOutput() {
        using var document = Create(); document.Recalculate();
        var report = document.AssessSave(Native());
        Assert.False(report.HasErrors);
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_FIELD_LOSS" && d.Location.EndsWith("/TotalSlackMinutes", StringComparison.Ordinal));
        using var output = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => document.Save(output, Native()));
        document.Save(output, Native(policy: OfficeConversionLossPolicy.Allow));
        using var loaded = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(document.Tasks.GetByUid(2).Finish, loaded.Tasks.GetByUid(2).Finish);
    }

    [Fact]
    public void UnmappedAssignmentCachesUseTheExplicitOmissionPolicy() {
        using var document = Create(); var assignment = document.Assignments.Single();
        assignment.PercentWorkComplete = 50; assignment.ActualStart = assignment.Start; assignment.ActualFinish = assignment.Finish;
        var report = document.AssessSave(Native()); Assert.False(report.HasErrors);
        foreach (string field in new[] { "PercentWorkComplete", "ActualStart", "ActualFinish" })
            Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_FIELD_LOSS" && d.Location.EndsWith("/" + field, StringComparison.Ordinal));
        using var output = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => document.Save(output, Native()));
        document.Save(output, Native(policy: OfficeConversionLossPolicy.Allow));
    }

    internal static ProjectDocument Create() {
        var document = ProjectDocument.Create(); document.Name = "Native project"; document.Title = "Native title";
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.ScheduleFromStart = true;
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var summary = document.Tasks.AddSummary("Delivery"); var task = summary.Children.Add("Design café / Łódź / 日本語");
        summary.IsManual = false; summary.IsActive = true; task.IsManual = false; task.IsActive = true;
        task.Duration = ProjectDuration.WorkingDays(1); task.Start = document.Settings.StartDate; task.Finish = task.Start.Value.AddHours(9);
        var resource = document.Resources.AddWork("Engineer"); resource.Calendar = document.Calendars.Add("Engineer", document.Calendar);
        resource.StandardRate = 100; resource.MaxUnits = ProjectUnits.Percent(100);
        var assignment = document.Assignments.Add(task, resource); assignment.Units = ProjectUnits.Percent(100); assignment.Work = ProjectWork.Hours(8);
        assignment.Start = task.Start; assignment.Finish = task.Finish;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Start = task.Start; baseline.Finish = task.Finish;
        baseline.Duration = task.Duration; baseline.Work = ProjectWork.Hours(8); baseline.Cost = 800;
        var definition = document.CustomFields.Add(); definition.FieldId = "188743731"; definition.FieldName = "Text1"; definition.Alias = "Work area";
        var value = task.CustomFields.Add(); value.FieldId = definition.FieldId; value.Value = "Independent authoring";
        var exception = document.Calendar.Exceptions.Add(); exception.Name = "Maintenance";
        exception.FromDate = new DateTime(2026, 10, 12); exception.ToDate = exception.FromDate; exception.IsWorking = false;
        return document;
    }
}
