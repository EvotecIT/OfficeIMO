using System.Collections.Generic;
using OfficeIMO.Core.Internal;
using System.Xml.Linq;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeBoundaryClosureTests {
    [Fact]
    public void NativeVariableValuesShareOneConfiguredBudgetAcrossTables() {
        byte[] source = File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp"));
        var compound = Compound(source);
        int count = new[] { "Cal", "Task", "Rsc", "Assn", "Cons" }
            .Sum(table => BitConverter.ToInt32(compound.Streams["   114/TBknd" + table + "/VarMeta"], 8));
        Assert.True(count > 1);
        using var accepted = ProjectDocument.Load(new MemoryStream(source), new ProjectLoadOptions { MaxNativeValues = count });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = count - 1 }));
        Assert.Contains("variable value budget", error.Message);
        Assert.Throws<ArgumentOutOfRangeException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = 0 }));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NativeVariableValueBudgetAppliesToEverySupportedGeneration(ProjectFileFormat format) {
        byte[] source = NewNative(format);
        using var accepted = ProjectDocument.Load(new MemoryStream(source));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = 1 }));
        Assert.Contains("variable value budget", error.Message);
    }

    [Fact]
    public void Mpp9SourceZeroVariableFieldsRoundTrip() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Tasks.GetByUid(2).Wbs = "DELIVERY.DESIGN";
        using var output = new MemoryStream(); document.Save(output, Native(ProjectFileFormat.Mpp9));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("DELIVERY.DESIGN", reopened.Tasks.GetByUid(2).Wbs);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryDoesNotEnterTheTypedModel(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        int[] expected = document.AllTasks.Select(task => task.Uid).ToArray(); Assert.DoesNotContain(0, expected);
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(expected, reopened.AllTasks.Select(task => task.Uid));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeIdentitiesDoNotMaterializeAbsentPublicGuids(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var explicitGuid = new Guid("57e42c89-4898-47e8-b184-d9130eb2d3bb");
        document.Tasks.GetByUid(2).Guid = explicitGuid;
        Assert.Null(document.Guid);
        Assert.All(document.AllTasks.Where(task => task.Uid != 2), task => Assert.Null(task.Guid));
        Assert.All(document.Resources, resource => Assert.Null(resource.Guid));
        Assert.All(document.Calendars, calendar => Assert.Null(calendar.Guid));
        Assert.All(document.Assignments, assignment => Assert.Null(assignment.Guid));

        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));

        Assert.Null(reopened.Guid);
        Assert.Equal(explicitGuid, reopened.Tasks.GetByUid(2).Guid);
        Assert.All(reopened.AllTasks.Where(task => task.Uid != 2), task => Assert.Null(task.Guid));
        Assert.All(reopened.Resources, resource => Assert.Null(resource.Guid));
        Assert.All(reopened.Calendars, calendar => Assert.Null(calendar.Guid));
        Assert.All(reopened.Assignments, assignment => Assert.Null(assignment.Guid));

        var later = reopened.Tasks.GetByUid(1).Children.Add("Later"); later.Duration = ProjectDuration.WorkingDays(1);
        var laterAssignment = reopened.Assignments.Add(later, reopened.Resources.GetByUid(1)); laterAssignment.Units = ProjectUnits.Percent(100);
        using var second = new MemoryStream(); reopened.Save(second, Native(format));
        using var twiceReopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Null(twiceReopened.Tasks.GetByUid(later.Uid).Guid);
        Assert.Null(twiceReopened.Assignments.GetByUid(laterAssignment.Uid).Guid);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryTracksAnExplicitProjectGuid(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Guid = new Guid("429757aa-d385-46b1-9161-8a6046c584a4");
        using var first = new MemoryStream(); document.Save(first, Native(format));
        document.Guid = new Guid("01c67d06-90f8-49a9-bfbd-c5b0665d84af");
        using var second = new MemoryStream(); document.Save(second, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.DoesNotContain(reopened.AllTasks, task => task.Uid == 0);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryTracksAProjectRename(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        using var first = new MemoryStream(); document.Save(first, Native(format));
        document.Name = "Renamed native project";
        using var second = new MemoryStream(); document.Save(second, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Equal("Renamed native project", reopened.Name);
        Assert.DoesNotContain(reopened.AllTasks, task => task.Uid == 0);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ExplicitProjectSummaryRemainsInTheTypedModel(ProjectFileFormat format) {
        using var document = ExplicitSummary();
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("Explicit project summary", reopened.Tasks.GetByUid(0).Name);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ExplicitNativeProjectSummaryDeletionIsRejected(ProjectFileFormat format) {
        using var source = ExplicitSummary(); using var output = new MemoryStream(); source.Save(output, Native(format));
        using var document = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.True(document.Tasks.Remove(document.Tasks.GetByUid(0)));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_PROJECT_SUMMARY_REQUIRED"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ReservedNativeResourceDeletionIsRejected(ProjectFileFormat format) {
        using var source = ProjectNativeAuthoringTests.Create(); using var output = new MemoryStream(); source.Save(output, Native(format));
        using var document = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.True(document.Resources.Remove(document.Resources.GetByUid(0)));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_RESERVED_RESOURCE_REQUIRED"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NativeCalendarKindUsesItsBaseReferenceWhenTheFlagIsAbsent(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var resource = document.Resources.First(item => item.Uid > 0); var derived = resource.Calendar!;
        Assert.NotNull(derived.BaseCalendar); derived.IsBaseCalendar = null;
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var copied = reopened.Resources.GetByUid(resource.Uid).Calendar!;
        Assert.False(copied.IsBaseCalendar); Assert.NotNull(copied.BaseCalendar);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void MissingNativeResourceTypeIsRejectedExplicitly(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var resource = document.Resources.AddWork("Untyped resource"); resource.Type = null;
        var options = Native(format);
        var report = document.AssessSave(options);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_RESOURCE_TYPE"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        using var output = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(output, options));
        Assert.Empty(output.ToArray()); Assert.Null(resource.Type);
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
    public void NewNativeTasksKeepAbsentStoredValuesAndReportBooleanDefaults(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        foreach (var item in document.AllTasks) { item.IsManual = null; item.IsActive = null; }
        var task = document.Tasks.GetByUid(2); var assignment = document.Assignments.Single();
        Assert.Null(task.Priority); Assert.Null(task.Type); Assert.Null(task.ConstraintType);
        Assert.Null(task.IsManual); Assert.Null(task.IsActive); Assert.Null(task.RemainingDuration);
        Assert.NotNull(task.Duration); Assert.NotNull(assignment.Work); Assert.Null(assignment.RemainingWork);

        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_TASK_DEFAULT"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/IsManual", StringComparison.Ordinal));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_TASK_DEFAULT"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/IsActive", StringComparison.Ordinal)
            && diagnostic.Message.EndsWith("to active.", StringComparison.Ordinal));
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        task = reopened.Tasks.GetByUid(2); assignment = reopened.Assignments.Single();
        Assert.Null(task.Priority); Assert.Null(task.Type); Assert.Null(task.ConstraintType);
        Assert.False(task.IsManual); Assert.True(task.IsActive); Assert.Null(task.RemainingDuration);
        Assert.NotNull(task.Duration); Assert.NotNull(assignment.Work); Assert.Null(assignment.RemainingWork);
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
    public void ClearingNativeTaskBooleanDefaultsReportsTheirNormalization(ProjectFileFormat format) {
        using var source = ProjectNativeAuthoringTests.Create();
        using var initial = new MemoryStream(); source.Save(initial, Native(format));
        using var document = ProjectDocument.Load(new MemoryStream(initial.ToArray()));
        var task = document.Tasks.GetByUid(2); task.IsManual = null; task.IsActive = null;
        var strict = new ProjectSaveOptions { Format = format };
        var report = document.AssessSave(strict);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_TASK_DEFAULT"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/IsManual", StringComparison.Ordinal));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_TASK_DEFAULT"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/IsActive", StringComparison.Ordinal)
            && diagnostic.Message.EndsWith(format == ProjectFileFormat.Mpp14 || format == ProjectFileFormat.Mpt14
                ? "to inactive." : "to active.", StringComparison.Ordinal));
        Assert.Throws<InvalidOperationException>(() => document.Save(new MemoryStream(), strict));

        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.False(reopened.Tasks.GetByUid(2).IsManual);
        Assert.Equal(format == ProjectFileFormat.Mpp14 || format == ProjectFileFormat.Mpt14
            ? false : true, reopened.Tasks.GetByUid(2).IsActive);
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
    public void NativeCalendarRangesReportTimeNormalization(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var exception = document.Calendar!.Exceptions[0];
        exception.FromDate = exception.FromDate!.Value.AddHours(8);
        exception.ToDate = exception.FromDate.Value.AddHours(4);

        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_CALENDAR_DATE_NORMALIZATION"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/FromDate", StringComparison.Ordinal));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_CALENDAR_DATE_NORMALIZATION"
            && diagnostic.RepresentsLoss && diagnostic.Location.EndsWith("/ToDate", StringComparison.Ordinal));

        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var retained = reopened.Calendar!.Exceptions[0];
        Assert.Equal(TimeSpan.Zero, retained.FromDate!.Value.TimeOfDay);
        Assert.Equal(new TimeSpan(23, 59, 0), retained.ToDate!.Value.TimeOfDay);
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
    public void NativeNullableEntityFlagsReportTheirGenerationSpecificDefaults(ProjectFileFormat format) {
        bool legacy8 = format == ProjectFileFormat.Mpp8 || format == ProjectFileFormat.Mpt8;
        foreach (string field in new[] { "TaskIsNull", "ResourceIsNull", "IsMilestone", "IsCritical", "EffortDriven" }) {
            using var document = ProjectNativeAuthoringTests.Create();
            var task = document.Tasks.GetByUid(2); var resource = document.Resources.Single(item => item.Uid > 0);
            switch (field) {
                case "TaskIsNull": task.IsNull = null; break;
                case "ResourceIsNull": resource.IsNull = null; break;
                case "IsMilestone": task.IsMilestone = null; break;
                case "IsCritical": task.IsCritical = null; break;
                case "EffortDriven": task.EffortDriven = null; break;
            }
            bool normalizes = field.EndsWith("IsNull", StringComparison.Ordinal) || !legacy8;
            string location = field == "TaskIsNull" || field == "ResourceIsNull" ? "/IsNull" : "/" + field;
            var report = document.AssessSave(new ProjectSaveOptions { Format = format });
            Assert.Equal(normalizes, report.Diagnostics.Any(d => d.RepresentsLoss && d.Location.EndsWith(location, StringComparison.Ordinal)));

            using var output = new MemoryStream(); document.Save(output, Native(format));
            using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
            bool? actual = field switch {
                "TaskIsNull" => reopened.Tasks.GetByUid(2).IsNull,
                "ResourceIsNull" => reopened.Resources.Single(item => item.Uid > 0).IsNull,
                "IsMilestone" => reopened.Tasks.GetByUid(2).IsMilestone,
                "IsCritical" => reopened.Tasks.GetByUid(2).IsCritical,
                _ => reopened.Tasks.GetByUid(2).EffortDriven
            };
            Assert.Equal(normalizes ? false : (bool?)null, actual);
        }
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ResourceCalendarWithoutBaseOrExplicitDerivedKindIsRejected(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var calendar = document.Resources.First(item => item.Uid > 0).Calendar!;
        calendar.BaseCalendar = null; calendar.IsBaseCalendar = null;
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_RESOURCE_CALENDAR"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void CrossProjectDependenciesAreRejectedBeforeNativeSerialization(ProjectFileFormat format) {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_CROSS_PROJECT_DEPENDENCY"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void MissingNativeDependencyTypeAndLagReportTheirDefaults(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var predecessor = document.Tasks.GetByUid(2); var successor = document.Tasks.Add("Successor");
        var dependency = document.Dependencies.Add(predecessor, successor); dependency.Type = null; dependency.Lag = null;
        var report = document.AssessSave(new ProjectSaveOptions { Format = format });
        Assert.Equal(2, report.Diagnostics.Count(diagnostic => diagnostic.Code == "PROJECT_NATIVE_DEPENDENCY_DEFAULT"
            && diagnostic.RepresentsLoss));
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var normalized = Assert.Single(reopened.Dependencies);
        Assert.Equal(ProjectDependencyType.FinishToStart, normalized.Type);
        Assert.True(normalized.Lag.HasValue); Assert.Equal(0m, normalized.Lag.Value.Value);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void MissingNativeCustomDurationFormatReportsItsDefault(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create(); var task = document.Tasks.GetByUid(2);
        var value = task.CustomFields.Add(); value.FieldId = "188743783"; value.Value = "PT1H0M0S";
        var report = document.AssessSave(new ProjectSaveOptions { Format = format });
        var diagnostic = Assert.Single(report.Diagnostics, item => item.Code == "PROJECT_NATIVE_CUSTOM_DURATION_DEFAULT");
        Assert.True(diagnostic.RepresentsLoss); Assert.EndsWith("/DurationFormat", diagnostic.Location);
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(7, reopened.Tasks.GetByUid(2).CustomFields.Single(item => item.FieldId == value.FieldId).DurationFormat);
    }

    [Fact]
    public void UnmappedNativeVariableFieldsFailBeforeTheirDataOffsetsAreRead() {
        byte[] source = File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp"));
        var compound = Compound(source);
        const string path = "   114/TBkndTask/VarMeta";
        byte[] metadata = (byte[])compound.Streams[path].Clone();
        Assert.True(BitConverter.ToInt32(metadata, 8) > 0);
        Buffer.BlockCopy(BitConverter.GetBytes(int.MaxValue), 0, metadata, 28, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, metadata, 32, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [path] = metadata });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(corrupt)));
        Assert.Contains("storage mapping", error.Message);
    }

    [Fact]
    public void FixedNativeFieldsCannotMasqueradeAsVariableStorage() {
        byte[] source = NewNative(ProjectFileFormat.Mpp14);
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default);
        var table = new ProjectNativeTable(compound, "TBkndTask", properties[0x03000014], properties[0x00020014],
            int.MaxValue, default, profile);
        uint fixedField = table.Fields.Values.First(field => field.Source == 10).Id;
        string path = profile.DataRoot + "/TBkndTask/VarMeta";
        byte[] metadata = (byte[])compound.Streams[path].Clone(); Assert.True(BitConverter.ToInt32(metadata, 8) > 0);
        Buffer.BlockCopy(BitConverter.GetBytes(int.MaxValue), 0, metadata, 28, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(fixedField), 0, metadata, 32, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [path] = metadata });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(corrupt)));
        Assert.Contains("qualified variable storage mapping", error.Message);
    }

    [Fact]
    public void NegativeNativeAssignmentRecordsDoNotEnterTheTypedOrXmlModels() {
        byte[] source = NewNative(ProjectFileFormat.Mpp14);
        byte[] corrupt = RewriteFixedUid(source, "Assn", 0x17, 0x0f400000);
        using var document = ProjectDocument.Load(new MemoryStream(corrupt));
        Assert.Empty(document.Assignments); Assert.NotEmpty(document.AllTasks); Assert.NotEmpty(document.Resources);
        using var xml = new MemoryStream();
        document.Save(xml, new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(xml.ToArray())); Assert.Empty(reopened.Assignments);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NativeTaskUidAndOutlineZeroMustIdentifyTheSameSummary(ProjectFileFormat format) {
        foreach (var mismatch in new[] { (Uid: 1, Level: 0), (Uid: 0, Level: 1) }) {
            byte[] source = RewriteFixedTaskField(NewNative(format), mismatch.Uid, 0x0b4000f9, mismatch.Level);
            var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
            Assert.Contains("outline level zero", error.Message);
        }
    }

    [Fact]
    public void NegativeNativeDependencyRecordsDoNotEnterTheTypedOrXmlModels() {
        using var sourceDocument = ProjectNativeAuthoringTests.Create();
        var predecessor = sourceDocument.Tasks.GetByUid(2); var successor = sourceDocument.Tasks.Add("Successor");
        sourceDocument.Dependencies.Add(predecessor, successor);
        using var source = new MemoryStream(); sourceDocument.Save(source, Native(ProjectFileFormat.Mpp14));
        byte[] corrupt = RewriteFixedUid(source.ToArray(), "Cons", 0x18, 0x0e400000);
        using var document = ProjectDocument.Load(new MemoryStream(corrupt)); Assert.Empty(document.Dependencies);
        using var xml = new MemoryStream();
        document.Save(xml, new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(xml.ToArray())); Assert.Empty(reopened.Dependencies);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NewNativeOutputDoesNotLeakStagingProperties(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Name = null; document.Title = null; document.Settings.ScheduleFromStart = null;
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Name); Assert.Null(reopened.Title); Assert.Null(reopened.Settings.ScheduleFromStart);
    }

    [Fact]
    public void MissingNativeSchedulingDirectionRemainsAbsentThroughEditsAndXmlConversion() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14), properties => properties.Remove(0x02400004));
        using var document = ProjectDocument.Load(new MemoryStream(source));
        Assert.Null(document.Settings.ScheduleFromStart);
        Assert.DoesNotContain("ScheduleFromStart", document.ToXml(new ProjectSaveOptions {
            Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow
        }));
        document.Tasks.First(task => task.Uid > 0).Name = "Edited without direction";
        using var output = new MemoryStream(); document.Save(output, Native(ProjectFileFormat.Mpp14));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Settings.ScheduleFromStart);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void MissingNativeDefaultClocksRemainAbsentThroughEditsAndXmlConversion(ProjectFileFormat format) {
        byte[] source = RewriteProperties(NewNative(format), properties => {
            properties.Remove(0x0240001c);
            properties.Remove(0x02400021);
        });
        using var document = ProjectDocument.Load(new MemoryStream(source));
        Assert.Null(document.Settings.DefaultStartTime);
        Assert.Null(document.Settings.DefaultFinishTime);
        string xml = document.ToXml(new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.DoesNotContain("DefaultStartTime", xml);
        Assert.DoesNotContain("DefaultFinishTime", xml);
        document.Tasks.First(task => task.Uid > 0).Name = "Edited without default clocks";
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Settings.DefaultStartTime);
        Assert.Null(reopened.Settings.DefaultFinishTime);
    }

    [Fact]
    public void UnknownNativeSchedulingDirectionIsRejected() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14),
            properties => properties.Set(0x02400004, 2, BitConverter.GetBytes((short)2)));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
        Assert.Contains("scheduling direction", error.Message);
    }

    [Fact]
    public void WrongWidthNativeSchedulingDirectionIsRejected() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14),
            properties => properties.Set(0x02400004, 4, BitConverter.GetBytes(65536)));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
        Assert.Contains("invalid width", error.Message);
    }

    private static ProjectSaveOptions Native(ProjectFileFormat format) => new() {
        Format = format, LossPolicy = OfficeConversionLossPolicy.Allow
    };

    private static ProjectDocument ExplicitSummary() {
        using var seed = ProjectNativeAuthoringTests.Create();
        foreach (var task in seed.AllTasks.Where(item => item.Duration.HasValue && !item.RemainingDuration.HasValue)) task.RemainingDuration = task.Duration;
        var xml = XDocument.Parse(seed.ToXml());
        XNamespace ns = XmlContracts.Ns; var tasks = xml.Root!.Element(ns + "Tasks")!;
        tasks.AddFirst(new XElement(ns + "Task", new XElement(ns + "UID", 0), new XElement(ns + "ID", 0),
            new XElement(ns + "Name", "Explicit project summary"), new XElement(ns + "OutlineLevel", 0), new XElement(ns + "Summary", 1)));
        return ProjectDocument.Parse(xml.ToString(SaveOptions.DisableFormatting));
    }

    private static byte[] NewNative(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        using var output = new MemoryStream(); document.Save(output, Native(format)); return output.ToArray();
    }

    private static byte[] RewriteProperties(byte[] source, Action<ProjectNativePropertySet> change) {
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = new ProjectNativePropertySet(compound.Streams[profile.Properties], default, profile == ProjectNativeProfile.Mpp8);
        change(properties);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> {
            [profile.Properties] = properties.Serialize(source.Length * 2L, default)
        });
    }

    private static byte[] RewriteFixedUid(byte[] source, string tableName, uint tableId, uint uidField) {
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default);
        var table = new ProjectNativeTable(compound, "TBknd" + tableName, properties[0x03000000 | tableId],
            properties[0x00020000 | tableId], int.MaxValue, default, profile);
        var record = Assert.Single(table.Records); var field = table.Fields[uidField]; Assert.Equal(10, field.Source);
        string prefix = profile.DataRoot + "/TBknd" + tableName + "/";
        byte[] metadata = compound.Streams[prefix + "FixedMeta"];
        byte[] data = (byte[])compound.Streams[prefix + "FixedData"].Clone();
        int recordOffset = BitConverter.ToInt32(metadata, 16 + record.MetadataIndex * table.MetadataWidth + 4);
        Buffer.BlockCopy(BitConverter.GetBytes(-1), 0, data, recordOffset + field.Offset, 4);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [prefix + "FixedData"] = data });
    }

    private static byte[] RewriteFixedTaskField(byte[] source, int uid, uint fieldId, int value) {
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default, profile == ProjectNativeProfile.Mpp8);
        ProjectNativeValue? legacyDescriptor = profile == ProjectNativeProfile.Mpp8 ? properties[0x02000001] : (ProjectNativeValue?)null;
        ProjectNativeTable table = legacyDescriptor.HasValue
            ? new ProjectNativeTable(compound, "TBkndTask", legacyDescriptor.Value, int.MaxValue, default)
            : new ProjectNativeTable(compound, "TBkndTask", properties[0x03000014],
                properties.TryGetValue(0x00020014, out var extended) ? extended : (ProjectNativeValue?)null, int.MaxValue, default, profile);
        var record = table.Records.Single(item => item.Integer(0x0b400056) == uid);
        var field = table.Fields[fieldId]; Assert.Equal(10, field.Source); Assert.False(field.Secondary);
        string prefix = profile.DataRoot + "/TBkndTask/";
        string dataName = legacyDescriptor.HasValue ? "FixFix   0" : "FixedData";
        byte[] data = (byte[])compound.Streams[prefix + dataName].Clone();
        int recordOffset;
        if (legacyDescriptor.HasValue) {
            var layout = ProjectNativeProperties.Read(legacyDescriptor.Value.Copy(), default);
            recordOffset = record.MetadataIndex * layout[5].Int32();
        } else {
            byte[] metadata = compound.Streams[prefix + "FixedMeta"];
            recordOffset = BitConverter.ToInt32(metadata, 16 + record.MetadataIndex * table.MetadataWidth + 4);
        }
        byte[] encoded = field.Size == 2 ? BitConverter.GetBytes(checked((short)value)) : BitConverter.GetBytes(value);
        Assert.Equal(field.Size, encoded.Length); Buffer.BlockCopy(encoded, 0, data, recordOffset + field.Offset, encoded.Length);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [prefix + dataName] = data });
    }

    private static OfficeCompoundFile Compound(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out var error), error);
        return compound!;
    }
}
