namespace OfficeIMO.Project.Tests;

public sealed class ProjectLegacyLifecycleRegressionTests {
    private static ProjectSaveOptions Options(ProjectFileFormat format) => new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };

    [Fact]
    public void AmbiguousInputDependenciesRemainOpaqueUntilExplicitSafeConversion() {
        const string text = "MPX-Fixture-4.0-ANSI\r\n61-90-1-74\r\n70-1-First\r\n70-2-Second-\"1FS-0.5d\"\r\n";
        byte[] input = Encoding.ASCII.GetBytes(text);
        using var document = ProjectDocument.Load(new MemoryStream(input));
        Assert.Empty(document.Dependencies); Assert.Contains(document.ReadDiagnostics, d => d.Code == "PROJECT_MPX_UNMODELED");
        using var retained = new MemoryStream(); document.Save(retained); Assert.Equal(input, retained.ToArray());
        document.Tasks[0].Name = "Edited";
        using var rejected = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(rejected, Options(ProjectFileFormat.Mpx4))); Assert.Empty(rejected.ToArray());
    }

    [Theory]
    [InlineData(100)]
    [InlineData(101)]
    public void MpxAssignmentLimitAppliesSeparatelyToEachTask(int resourceCount) {
        var text = new StringBuilder("MPX,Fixture,4.0,ANSI\r\n41,40,1\r\n");
        for (int resource = 1; resource <= resourceCount; resource++) text.Append("50,").Append(resource).Append(",Resource ").Append(resource).Append("\r\n");
        text.Append("61,90,1\r\n");
        for (int task = 1; task <= 2; task++) {
            text.Append("70,").Append(task).Append(",Task ").Append(task).Append("\r\n");
            for (int resource = 1; resource <= resourceCount; resource++) text.Append("75,").Append(resource).Append(",1,8h\r\n");
        }
        using var bytes = new MemoryStream(Encoding.ASCII.GetBytes(text.ToString()));
        if (resourceCount > 100) Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(bytes));
        else { using var document = ProjectDocument.Load(bytes); Assert.Equal(200, document.Assignments.Count); }
    }

    [Theory]
    [InlineData(',')]
    [InlineData(';')]
    [InlineData('|')]
    public void MpxPredecessorListsRetainSignedFractionalLags(char separator) {
        using var document = ProjectDocument.Create();
        var first = document.Tasks.Add("First"); var second = document.Tasks.Add("Second"); var last = document.Tasks.Add("Last");
        document.Dependencies.Add(first, last).Lag = ProjectDuration.WorkingDays(-0.5m);
        document.Dependencies.Add(second, last).LagPercent = 25;
        var options = Options(ProjectFileFormat.Mpx4); options.MpxSeparator = separator;
        using var bytes = new MemoryStream(); document.Save(bytes, options);
        using var read = ProjectDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(2, read.Dependencies.Count);
        Assert.Equal(ProjectDuration.WorkingDays(-0.5m), read.Dependencies.Single(d => d.Predecessor!.Uid == first.Uid).Lag);
        Assert.Equal(25m, read.Dependencies.Single(d => d.Predecessor!.Uid == second.Uid).LagPercent);
    }

    [Theory]
    [InlineData('-')]
    [InlineData('+')]
    [InlineData('.')]
    [InlineData('%')]
    public void MpxAmbiguousSeparatorsFailBeforeOutput(char separator) {
        using var document = ProjectDocument.Create();
        var first = document.Tasks.Add("First"); var second = document.Tasks.Add("Second");
        document.Dependencies.Add(first, second).Lag = ProjectDuration.WorkingDays(-0.5m);
        var options = Options(ProjectFileFormat.Mpx4); options.MpxSeparator = separator;
        using var bytes = new MemoryStream(); bytes.WriteByte(123);
        Assert.Throws<InvalidDataException>(() => document.Save(bytes, options));
        Assert.Equal(new byte[] { 123 }, bytes.ToArray());
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8, false)]
    [InlineData(ProjectFileFormat.Mpp9, false)]
    [InlineData(ProjectFileFormat.Mpp8, true)]
    [InlineData(ProjectFileFormat.Mpp9, true)]
    public void LegacyDerivedCalendarsDropRemovedAncestorProjections(ProjectFileFormat format, bool rebind) {
        using var document = ProjectNativeAuthoringTests.Create();
        var day = new DateTime(2026, 10, 5);
        var week = document.Calendar!.WorkWeeks.Add(); week.FromDate = day; week.ToDate = day;
        week.SetWorkingDay(DayOfWeek.Monday, ProjectWorkingTime.Hours(9, 12));
        var derived = document.Resources.Single().Calendar!;
        using var first = new MemoryStream(); document.Save(first, Options(format));
        using (var initial = ProjectDocument.Load(new MemoryStream(first.ToArray())))
            Assert.Equal(180m, initial.Calendars.GetByUid(derived.Uid).WorkingMinutesBetween(day, day.AddDays(1)));
        if (rebind) derived.BaseCalendar = document.Calendars.AddStandardWorkingWeek("Replacement");
        else document.Calendar.WorkWeeks.Remove(week);
        using var second = new MemoryStream(); document.Save(second, Options(format));
        using var read = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Equal(480m, read.Calendars.GetByUid(derived.Uid).WorkingMinutesBetween(day, day.AddDays(1)));
    }

    [Fact]
    public void MpxResourceCalendarsDoNotConsumeBaseCalendarCapacity() {
        var text = new StringBuilder("MPX,Fixture,4.0,ANSI\r\n20,Standard,0,1,1,1,1,1,0\r\n");
        for (int day = 2; day <= 6; day++) text.Append("25,").Append(day).Append(",08:00,12:00,13:00,17:00\r\n");
        text.Append("41,40,1\r\n");
        for (int i = 1; i <= 250; i++) text.Append("50,").Append(i).Append(",Resource ").Append(i).Append("\r\n55,Standard,2,2,2,2,2,2,2\r\n");
        using var document = ProjectDocument.Load(new MemoryStream(Encoding.ASCII.GetBytes(text.ToString())));
        Assert.Equal(251, document.Calendars.Count); document.Resources[0].Name = "Renamed";
        using var bytes = new MemoryStream(); document.Save(bytes, Options(ProjectFileFormat.Mpx4));
        using var read = ProjectDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(251, read.Calendars.Count); Assert.Equal("Renamed", read.Resources[0].Name);
        Assert.All(read.Resources, resource => Assert.Equal(read.Calendar, resource.Calendar!.BaseCalendar));
    }
}
