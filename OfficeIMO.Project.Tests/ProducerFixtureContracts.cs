namespace OfficeIMO.Project.Tests;

public class ProducerFixtureContracts {
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Project2024", name + ".xml");

    [Fact]
    public void CostResourceAmountsRemainStoredAndExposeTheApplicationImportLimit() {
        using var project = ProjectDocument.Load(Fixture("resources"));
        var travel = project.Resources.Single(r => r.Name == "Travel");
        Assert.Equal(ProjectResourceType.Cost, travel.Type);
        Assert.Equal(300m, travel.Cost);
        var assignment = project.Assignments.Single(a => a.Resource == travel);
        Assert.Equal(300m, assignment.Cost);
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_COST_RESOURCE_IMPORT");
        assignment.Cost = 325; assignment.RemainingCost = 325;
        using var copy = ProjectDocument.Parse(project.ToXml());
        Assert.Equal(325m, copy.Assignments.Single(a => a.Resource?.Name == "Travel").Cost);
    }

    [Fact]
    public void CalendarDateEditsAndRemovalKeepLegacyAndModernRepresentationsTogether() {
        using var project = ProjectDocument.Load(Fixture("calendars"));
        var calendar = project.Calendars.Single(c => c.Name == "Workshop");
        var exception = calendar.Exceptions.Single();
        exception.FromDate = new DateTime(2026, 10, 13); exception.ToDate = exception.FromDate;
        var xml = System.Xml.Linq.XDocument.Parse(project.ToXml());
        var node = xml.Descendants(System.Xml.Linq.XName.Get("Calendar", XmlContracts.Ns))
            .Single(c => (string?)c.Element(System.Xml.Linq.XName.Get("Name", XmlContracts.Ns)) == "Workshop");
        Assert.Equal(2, node.Descendants(System.Xml.Linq.XName.Get("FromDate", XmlContracts.Ns)).Count(d => d.Value.StartsWith("2026-10-13", StringComparison.Ordinal)));
        calendar.Exceptions.Remove(exception);
        using var copy = ProjectDocument.Parse(project.ToXml(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var result = copy.Calendars.Single(c => c.Name == "Workshop");
        Assert.Empty(result.Exceptions);
        Assert.DoesNotContain(result.WeekDays, d => d.Day == null);
    }

    [Theory]
    [InlineData("empty", 1)]
    [InlineData("delivery", 5)]
    [InlineData("calendars", 5)]
    [InlineData("actuals", 5)]
    [InlineData("relationships", 11)]
    [InlineData("resources", 5)]
    [InlineData("custom-fields", 5)]
    public void Project2024FixturesValidateAndPreserveUnchangedBytes(string name, int taskCount) {
        using var document = ProjectDocument.Load(Fixture(name));
        document.Validate().ThrowIfErrors();
        Assert.Equal(14, document.SourceSaveVersion);
        Assert.Equal(taskCount, document.AllTasks.Count());
        using var output = new MemoryStream();
        document.Save(output);
        Assert.Equal(File.ReadAllBytes(Fixture(name)), output.ToArray());
        if (name != "empty") {
            var build = document.AllTasks.Single(t => t.Name == "Build");
            Assert.Equal("Synthetic notes: café / Łódź / 日本語", build.Notes);
            Assert.Equal(5000m, build.Baselines.Single(b => b.Number == 0).Cost);
            Assert.Contains(build.CustomFields, f => f.Value == "Platform");
        }
    }

    [Fact]
    public void ActualsCostsAndRatesMatchTheProducerObjectModel() {
        using var document = ProjectDocument.Load(Fixture("actuals"));
        var build = document.AllTasks.Single(t => t.Name == "Build");
        Assert.Equal(40, build.PercentComplete);
        Assert.Equal(960m, build.ActualWork!.Value.Minutes);
        Assert.Equal(1440m, build.RemainingWork!.Value.Minutes);
        Assert.Equal(2073.95m, build.ActualCost);
        Assert.Equal(120.50m, build.FixedCost);
        Assert.Equal(5146.25m, build.Cost);
        var engineer = document.Resources.Single(r => r.Name == "Engineer");
        Assert.Equal(125m, engineer.StandardRate);
        Assert.Equal(25.75m, engineer.CostPerUse);
        Assert.NotEmpty(document.Assignments.SelectMany(a => a.TimephasedData));
    }

    [Fact]
    public void CalendarExceptionsAndReferenceEditsUseProducerIdentities() {
        using var document = ProjectDocument.Load(Fixture("calendars"));
        var calendar = document.Calendars.Single(c => c.Name == "Workshop");
        Assert.False(calendar.WeekDays.Single(d => d.Day == DayOfWeek.Friday).IsWorking);
        var exception = Assert.Single(calendar.Exceptions);
        Assert.Equal("Maintenance", exception.Name);
        Assert.Equal(new DateTime(2026, 10, 12), exception.FromDate);
        var build = document.AllTasks.Single(t => t.Name == "Build");
        Assert.Same(calendar, build.Calendar);
        build.Calendar = null;
        document.Validate().ThrowIfErrors();
        using var copy = ProjectDocument.Parse(document.ToXml());
        Assert.Null(copy.Tasks.GetByUid(build.Uid).Calendar);
    }

    [Fact]
    public void AllDependencyKindsAndLagScalesMatchProject2024() {
        using var document = ProjectDocument.Load(Fixture("relationships"));
        foreach (var pair in new[] {
            ("Dependency FS", ProjectDependencyType.FinishToStart),
            ("Dependency SS", ProjectDependencyType.StartToStart),
            ("Dependency FF", ProjectDependencyType.FinishToFinish),
            ("Dependency SF", ProjectDependencyType.StartToFinish) }) {
            var link = document.Dependencies.Single(d => d.Successor.Name == pair.Item1);
            Assert.Equal(pair.Item2, link.Type);
            Assert.Equal(ProjectDuration.WorkingHours(2), link.Lag);
        }
        Assert.Equal(ProjectDuration.WorkingDays(-1), document.Dependencies.Single(d => d.Successor.Name == "Negative lag").Lag);
        Assert.Equal(50m, document.Dependencies.Single(d => d.Successor.Name == "Percent lag").LagPercent);
    }
}
