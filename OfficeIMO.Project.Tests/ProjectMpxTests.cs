using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectMpxTests {
    private static ProjectDocument Read(string text) => ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(text)));
    private static ProjectSaveOptions Options(bool allow = true) => new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4, LossPolicy = allow ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block };

    [Theory]
    [InlineData("  padded  ")]
    [InlineData("\tpadded\t")]
    [InlineData(" \"quoted\" ")]
    [InlineData(" \t ")]
    public void QuotedPayloadWhitespaceSurvivesReadingAndEditedRewrites(string value) {
        string quoted = " \t\"" + value.Replace("\"", "\"\"") + "\"\t ";
        string source = "MPX,Fixture,4.0,ANSI\r\n41,40,49,1\r\n50,1,1," + quoted
            + "\r\n61,90,98,1,14\r\n70,1,1," + quoted + "," + quoted + "\r\n";
        using var project = Read(source);
        Assert.Equal(value, project.Tasks[0].Name); Assert.Equal(value, project.Tasks[0].Notes); Assert.Equal(value, project.Resources[0].Name);
        project.Tasks[0].Name = value + " edited ";
        using var output = new MemoryStream(); project.Save(output, Options());
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(value + " edited ", reopened.Tasks[0].Name);
        Assert.Equal(value, reopened.Tasks[0].Notes); Assert.Equal(value, reopened.Resources[0].Name);
    }

    [Fact]
    public void UnquotedPaddingAndEmptyQuotedFieldsHaveDistinctFraming() {
        using var project = Read("MPX,Fixture,4.0,ANSI\r\n61,90,98,1,14\r\n70,1,1, \tTask\t , \t\"\" \t\r\n");
        Assert.Equal("Task", project.Tasks[0].Name); Assert.True(string.IsNullOrEmpty(project.Tasks[0].Notes));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpaqueAssignmentDelayBlocksScheduling(bool calculateAssignments) {
        const string source = "MPX,OfficeIMO fixture,4.0,ANSI\r\n41,40,49,1,42\r\n50,1,1,Engineer,100/h\r\n61,90,98,1,40\r\n70,1,1,Task,1d\r\n75,1,1,8h,,,,,,,,,1d,1\r\n";
        using var document = Read(source); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = calculateAssignments });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_MPX_SCHEDULING_PROFILE" && d.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void DeclaredLocaleControlsDatesCurrencyDurationAndQuotedNotes() {
        const string source = "MPX;OfficeIMO fixture;4.0;ANSI\r\n10;EUR;1;2;.;,\r\n11;2;0;1;7,5;37,5\r\n12;1;1;480;/;:;AM;PM\r\n" +
            "41;40;49;1;42\r\n50;4;91;Engineer;125,5/h\r\n61;90;98;1;40;50;51;30;14;74\r\n" +
            "70;3;81;\"Design; quoted \"\"name\"\"\";1,5d;05/10/2026 08:00;06/10/2026 12:00;EUR1.234,50;\"line 1\u007fline 2; notes\"\r\n" +
            "70;5;82;Build;2ed;07/10/2026 08:00;09/10/2026 08:00;;;81FS+0,5d\r\n75;4;0,5;15h;;3h;;125,5;;;;;;91\r\n";
        using var project = Read(source);
        var design = project.Tasks.GetByUid(81); var build = project.Tasks.GetByUid(82);
        Assert.Equal("Design; quoted \"name\"", design.Name);
        Assert.Equal("line 1\nline 2; notes", design.Notes);
        Assert.Equal(1234.5m, design.Cost); Assert.Equal(450, project.Settings.MinutesPerDay);
        Assert.Equal(new DateTime(2026, 10, 5, 8, 0, 0), design.Start);
        Assert.Equal(ProjectDuration.WorkingDays(1.5m), design.Duration);
        Assert.Equal(ProjectDuration.ElapsedDays(2), build.Duration);
        Assert.Equal(ProjectDuration.WorkingDays(0.5m), project.Dependencies.Single().Lag);
        Assert.Equal(81, project.Dependencies.Single().Predecessor!.Uid);
        Assert.Equal(91, project.Assignments.Single().Resource!.Uid);
        Assert.Equal(900m, project.Assignments.Single().Work!.Value.Minutes);
        Assert.Equal(125.5m, project.Resources.Single().StandardRate);
        using var output = new MemoryStream(); project.Save(output);
        Assert.Equal(Encoding.ASCII.GetBytes(source), output.ToArray());
    }

    [Theory]
    [InlineData(ProjectMpxEncoding.Windows1252, "café €")]
    [InlineData(ProjectMpxEncoding.Dos437, "café Ω")]
    [InlineData(ProjectMpxEncoding.Dos850, "café ø")]
    [InlineData(ProjectMpxEncoding.MacintoshRoman, "café ")]
    public void NewFilesUseExplicitCodePagesAndKeepIdsAcrossEdits(ProjectMpxEncoding encoding, string name) {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        var summary = project.Tasks.AddSummary(name); var first = summary.Children.Add("First"); var second = summary.Children.Add("Second");
        first.Duration = ProjectDuration.WorkingHours(3); first.Start = new DateTime(2026, 10, 5, 8, 0, 0); first.Finish = new DateTime(2026, 10, 5, 11, 0, 0);
        var resource = project.Resources.AddWork("Engineer"); resource.StandardRate = 125;
        var assignment = project.Assignments.Add(second, resource, ProjectUnits.Percent(50)); assignment.Work = ProjectWork.Hours(4); assignment.Cost = 500;
        project.Dependencies.Add(first, second).LagPercent = 50;
        var options = Options(); options.MpxEncoding = encoding; options.MpxSeparator = ';';
        using var output = new MemoryStream(); project.Save(output, options);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(name, reopened.Tasks[0].Name); Assert.Equal(summary.Uid, reopened.Tasks.GetByUid(second.Uid).Parent!.Uid);
        Assert.Equal(first.Start, reopened.Tasks.GetByUid(first.Uid).Start);
        Assert.Equal(50m, reopened.Dependencies.Single().LagPercent);
        Assert.Equal(240m, reopened.Assignments.Single().Work!.Value.Minutes);
        reopened.Tasks.GetByUid(second.Uid).MoveTo(null); reopened.Tasks.GetByUid(first.Uid).Name = name;
        using var edited = new MemoryStream(); reopened.Save(edited, Options());
        using var final = ProjectDocument.Load(new MemoryStream(edited.ToArray()));
        Assert.Null(final.Tasks.GetByUid(second.Uid).Parent); Assert.Equal(name, final.Tasks.GetByUid(first.Uid).Name);
        Assert.Equal(resource.Uid, final.Assignments.Single().Resource!.Uid);
    }

    [Fact]
    public void UnsupportedEncodingFailsBeforeReplacingDestination() {
        using var project = ProjectDocument.Create(); project.Tasks.Add("日本語");
        using var output = new MemoryStream(); output.WriteByte(123);
        Assert.Contains(project.AssessSave(Options()).Diagnostics, d => d.Code == "PROJECT_MPX_UNREPRESENTABLE" && d.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => project.Save(output, Options())); Assert.Equal(new byte[] { 123 }, output.ToArray());
    }

    [Fact]
    public void UnmodeledRecordsAreInertAndRequireLossAcceptanceOnlyOnRewrite() {
        const string source = "MPX,Fixture,4.0,ANSI\r\n61,90,98,1\r\n70,1,15,Task\r\n80,external-project.mpp\r\n81,untrusted-app,topic,item\r\n";
        using var project = Read(source);
        Assert.Contains(project.ReadDiagnostics, d => d.Code == "PROJECT_MPX_UNMODELED");
        using var unchanged = new MemoryStream(); project.Save(unchanged); Assert.Equal(Encoding.ASCII.GetBytes(source), unchanged.ToArray());
        project.Tasks[0].Name = "Changed";
        using var edited = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => project.Save(edited, Options(false))); Assert.Empty(edited.ToArray());
        project.Save(edited, Options());
        using var reopened = ProjectDocument.Load(new MemoryStream(edited.ToArray())); Assert.Equal("Changed", reopened.Tasks[0].Name);
    }

    [Fact]
    public void LossAssessmentSurvivesSaveAndPendingBatches() {
        using var project = ProjectDocument.Create(); var task = project.Tasks.Add("Task"); task.Priority = 551;
        using var output = new MemoryStream(); project.Save(output, Options());
        Assert.Contains(project.AssessSave(Options(false)).Diagnostics, d => d.Code == "PROJECT_MPX_PRIORITY");
        Assert.Throws<InvalidOperationException>(() => project.Save(new MemoryStream(), Options(false)));
        using (project.BeginUpdate()) {
            task.Name = "日本語";
            Assert.Contains(project.AssessSave(Options()).Diagnostics, d => d.Severity == ProjectDiagnosticSeverity.Error);
        }
    }

    [Theory]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n61,90,1\r\n70,1,\"unterminated")]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n61,90,1\r\n70,1,\"closed\"junk\r\n")]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n61,90,1\r\n70,1,a\r\n70,1,b\r\n")]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n71,orphan\r\n")]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n61,90,1,3\r\n70,1,Task,3\r\n")]
    [InlineData("MPX,Fixture,4.0,ANSI\r\n12,2\r\n61,90,1,50\r\n70,1,Task,invalid\r\n")]
    public void MalformedRecordsFailClosed(string text) => Assert.Throws<InvalidDataException>(() => Read(text));

    [Fact]
    public void MpxToXmlRequiresAnExplicitCurrencyIdentity() {
        using var project = Read("MPX,Fixture,4.0,ANSI\r\n61,90,1\r\n70,1,Task\r\n");
        Assert.Throws<InvalidDataException>(() => project.ToXml());
        project.Settings.CurrencyCode = "USD";
        using var xml = ProjectDocument.Parse(project.ToXml()); Assert.Equal("Task", xml.Tasks[0].Name);
    }

    [Fact]
    public void LiteralNaTextIsDistinctFromAnAbsentDate() {
        using var project = Read("MPX,Fixture,4.0,ANSI\r\n61,90,1,50,4\r\n70,1,NA,NA,NA\r\n");
        Assert.Equal("NA", project.Tasks[0].Name); Assert.Null(project.Tasks[0].Start);
        Assert.Equal("NA", project.Tasks[0].CustomFields.Single().Value);
    }

    [Fact]
    public void RepeatedResourceCalendarEditsDoNotAccumulateCalendars() {
        byte[] bytes;
        using (var project = ProjectDocument.Create()) {
            project.Calendar = project.Calendars.AddStandardWorkingWeek();
            var resource = project.Resources.AddWork("Engineer"); resource.Calendar = project.Calendar;
            using var output = new MemoryStream(); project.Save(output, Options()); bytes = output.ToArray();
        }
        for (int iteration = 0; iteration < 3; iteration++) {
            using var project = ProjectDocument.Load(new MemoryStream(bytes));
            Assert.Equal(2, project.Calendars.Count); Assert.Equal(project.Calendar, project.Resources[0].Calendar!.BaseCalendar);
            project.Resources[0].Name = "Engineer " + iteration;
            using var output = new MemoryStream(); project.Save(output, Options()); bytes = output.ToArray();
        }
    }

    [Fact]
    public void CustomCostsUseModelHundredthsAndStartFinishDatesRemainSeparate() {
        using var project = Read("MPX,Fixture,4.0,ANSI\r\n12,1,1,480,/,:\r\n61,90,1,30,36,60,61\r\n70,1,Task,1.5,1.5,Mon 05/10/26 08:00,06 October 2026 17:00\r\n");
        var task = project.Tasks[0];
        Assert.Equal(1.5m, task.Cost);
        Assert.Equal(150m, decimal.Parse(task.CustomFields.Single(f => f.FieldId == "188743786").Value!, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("2026-10-05T08:00:00", task.CustomFields.Single(f => f.FieldId == "188743732").Value);
        Assert.Equal("2026-10-06T17:00:00", task.CustomFields.Single(f => f.FieldId == "188743733").Value);
        task.Name = "Edited";
        using var output = new MemoryStream(); project.Save(output, Options());
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(task.CustomFields.Select(f => f.Value), reopened.Tasks[0].CustomFields.Select(f => f.Value));
    }

    [Fact]
    public void FractionalProgressReportsNormalizationAndPreservesUnchangedBytes() {
        const string input = "MPX;Fixture;4.0;ANSI\r\n10;$;1;2;.;,\r\n61;90;1;44\r\n70;1;Task;55,5%\r\n";
        using var project = Read(input); Assert.Equal(56, project.Tasks[0].PercentComplete);
        Assert.Contains(project.ReadDiagnostics, d => d.RepresentsLoss && d.Message.StartsWith("Fractional progress", StringComparison.Ordinal));
        using var output = new MemoryStream(); project.Save(output); Assert.Equal(Encoding.ASCII.GetBytes(input), output.ToArray());
        project.Tasks[0].Name = "Changed";
        Assert.Throws<InvalidOperationException>(() => project.Save(new MemoryStream(), Options(false)));
    }

    [Fact]
    public void CommentsAndPresentationDeclarationsSurviveAnEdit() {
        using var project = Read("MPX,Fixture,4.0,ANSI\r\n0,\"A comment, with a separator\"\r\n10,$,3,2,\",\",.\r\n12,1,0,480,/,:,AM,PM,20,9\r\n61,90,1\r\n70,1,Task\r\n");
        project.Tasks[0].Name = "Edited";
        using var output = new MemoryStream(); project.Save(output, Options());
        var records = ProjectMpxRecords.Read(output.ToArray(), new ProjectLoadOptions(), default).Records;
        Assert.Equal("A comment, with a separator", records.Single(r => r[0] == "0")[1]);
        Assert.Equal("3", records.Single(r => r[0] == "10")[2]);
        Assert.Equal("20", records.Single(r => r[0] == "12")[8]);
        Assert.Equal("9", records.Single(r => r[0] == "12")[9]);
    }

    [Fact]
    public void LimitsCancellationAndPrecisionFailBeforeDestinationChanges() {
        const string text = "MPX,Fixture,4.0,ANSI\r\n61,90,1\r\n70,1,First\r\n70,2,Second\r\n";
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(text)), new ProjectLoadOptions { MaxTasks = 1 }));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(text)), new ProjectLoadOptions { MaxElements = 5 }));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(text)), cancellationToken: cancelled.Token));
        using var project = ProjectDocument.Create(); var task = project.Tasks.Add("Task"); task.Start = new DateTime(2026, 10, 5, 8, 0, 1);
        using var output = new MemoryStream(); output.WriteByte(5);
        Assert.Contains(project.AssessSave(Options()).Diagnostics, d => d.Code == "PROJECT_MPX_DATE_PRECISION");
        Assert.Throws<InvalidDataException>(() => project.Save(output, Options())); Assert.Equal(new byte[] { 5 }, output.ToArray());
    }
}
