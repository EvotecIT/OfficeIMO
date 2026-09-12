namespace OfficeIMO.Project.Tests;

public sealed class ProjectDataTests {
    [Theory]
    [InlineData("0.0000000000000000000000000001", true)]
    [InlineData("0.000000000000000000000000000001", false)]
    [InlineData("-0.000000000000000000000000000001", false)]
    [InlineData("1.12345678901234567890123456789", false)]
    [InlineData("1.250000000000000000000000000000000", true)]
    public void DecimalImportRequiresExactRepresentation(string cost, bool accepted) {
        var fields = new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.Cost };
        var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] { new[] { "1", "Task", cost } });
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, fields.Select(f => new ProjectDataColumn(f, f.ToString())));
        if (!accepted) { Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped })); return; }
        using var document = ProjectDocument.ImportTables(new[] { mapped }).Document;
        Assert.Equal(decimal.Parse(cost, System.Globalization.CultureInfo.InvariantCulture), document.Tasks[0].Cost);
    }

    [Theory]
    [InlineData(DateTimeKind.Utc, false)]
    [InlineData(DateTimeKind.Local, false)]
    [InlineData(DateTimeKind.Utc, true)]
    [InlineData(DateTimeKind.Local, true)]
    public void ExportRejectsTimezoneBearingEntityDates(DateTimeKind kind, bool assignmentDate) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task");
        var date = new DateTime(2026, 10, 5, 8, 0, 0, kind);
        if (assignmentDate) document.Assignments.Add(task, document.Resources.AddWork("Engineer")).Start = date;
        else task.Start = date;
        Assert.Throws<InvalidDataException>(() => document.ExportTables(allowLossyProjection: true));
    }

    [Fact]
    public void FourTablesRoundTripIdentitiesOutlineUnitsAndWorkingWeek() {
        using var original = ProjectDocument.Create();
        original.Calendar = original.Calendars.AddStandardWorkingWeek();
        original.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var summary = original.Tasks.AddSummary("Plan");
        var task = summary.Children.Add("Zażółć, \"quote\"\nnew line"); task.Duration = ProjectDuration.WorkingHours(8);
        task.Calendar = original.Calendar; task.Work = ProjectWork.Hours(4);
        var resource = original.Resources.AddWork("Engineer"); resource.MaxUnits = ProjectUnits.Percent(50);
        var assignment = original.Assignments.Add(task, resource, ProjectUnits.Percent(50)); assignment.Work = ProjectWork.Hours(4);
        Assert.Throws<InvalidOperationException>(() => original.ExportTables());
        var export = original.ExportTables(allowLossyProjection: true); Assert.Equal(4, export.Tables.Count); Assert.NotEmpty(export.Notices);
        var renamed = export.WithColumns(ProjectDataKind.Resources, new[] { new ProjectDataColumn(ProjectDataField.Name, "Label"), new ProjectDataColumn(ProjectDataField.Uid, "Identity") });
        var renamedResources = renamed.Tables.Single(t => t.Kind == ProjectDataKind.Resources);
        Assert.Equal(new[] { "Label", "Identity" }, renamedResources.Table.Headers);
        Assert.Equal(resource.Name, renamedResources.Table.Rows[0][0]); Assert.Contains(renamed.Notices, n => n.Contains("omitted mapped fields"));
        var result = ProjectDocument.ImportTables(export.Tables, new ProjectDataImportOptions { Start = original.Settings.StartDate, CalendarUid = original.Calendar.Uid });
        using var imported = result.Document;
        Assert.Empty(result.Notices);
        var copy = imported.Tasks.GetByUid(task.Uid); Assert.Equal(task.Name, copy.Name); Assert.Equal(summary.Uid, copy.Parent!.Uid);
        Assert.Equal(480m, copy.Duration!.Value.Value); Assert.Equal(ProjectDurationUnit.Minute, copy.Duration.Value.Unit);
        Assert.Equal(.5m, imported.Resources.GetByUid(resource.Uid).MaxUnits!.Value.Value);
        Assert.Equal(assignment.Uid, Assert.Single(imported.Assignments).Uid);
        Assert.Equal(.5m, imported.Assignments[0].Units!.Value.Value);
        var schedule = imported.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        Assert.Equal(240m, schedule.Assignments.Single().Work.Minutes);
        Assert.Equal(new DateTime(2026, 10, 5, 17, 0, 0), schedule.Tasks.Single(t => t.TaskUid == task.Uid).Finish);
        Assert.NotEqual(task.Uid, imported.Tasks.Add("New task").Uid);
    }

    [Fact]
    public void TablesRoundTripAnExplicitlyUnassignedResource() {
        const string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\"><Tasks><Task><UID>1</UID><Name>Task</Name></Task></Tasks>"
            + "<Assignments><Assignment><UID>7</UID><TaskUID>1</TaskUID></Assignment></Assignments></Project>";
        using var source = ProjectDocument.Parse(xml);
        source.Validate().ThrowIfErrors();
        var exported = source.ExportTables(allowLossyProjection: true);
        var assignmentTable = exported.Tables.Single(t => t.Kind == ProjectDataKind.Assignments);
        int resourceColumn = Array.IndexOf(assignmentTable.Table.Headers.ToArray(), ProjectDataField.ResourceUid.ToString());
        Assert.Equal("-1", Assert.Single(assignmentTable.Table.Rows)[resourceColumn]);
        var result = ProjectDocument.ImportTables(exported.Tables);
        using var imported = result.Document;
        var assignment = Assert.Single(imported.Assignments);
        Assert.Equal(7, assignment.Uid);
        Assert.Equal(1, assignment.Task!.Uid);
        Assert.Null(assignment.Resource);
        Assert.Equal(-1, assignment.SourceResourceUid);
    }

    [Fact]
    public void ExplicitHeaderMappingPreservesSparseIdentitiesAndForwardParents() {
        var table = new ProjectDataTable(new[] { "Identifier", "Label", "Parent", "Duration", "Extra" }, new[] {
            new[] { "900", "Child", "42", "60", "ignored" }, new[] { "42", "Parent", "", "", "ignored" }
        });
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, new[] {
            new ProjectDataColumn(ProjectDataField.Uid, "Identifier"), new ProjectDataColumn(ProjectDataField.Name, "Label"),
            new ProjectDataColumn(ProjectDataField.ParentUid, "Parent"), new ProjectDataColumn(ProjectDataField.DurationMinutes, "Duration")
        });
        Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped }));
        var result = ProjectDocument.ImportTables(new[] { mapped }, new ProjectDataImportOptions { AllowUnmappedColumns = true });
        using var project = result.Document;
        Assert.Single(result.Notices); Assert.Equal(42, Assert.Single(project.Tasks).Uid);
        Assert.Equal(900, Assert.Single(project.Tasks[0].Children).Uid);
    }

    [Theory]
    [InlineData("1", "2", "2", "1")]
    [InlineData("1", "99", "2", "")]
    [InlineData("1", "", "1", "")]
    public void ConflictingOrCyclicIdentityRejectsTheEntireImport(string first, string parent, string second, string otherParent) {
        var table = new ProjectDataTable(new[] { "Uid", "Name", "ParentUid" }, new[] { new[] { first, "First", parent }, new[] { second, "Second", otherParent } });
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.ParentUid }.Select(f => new ProjectDataColumn(f, f.ToString())));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped }));
    }

    [Fact]
    public void ImportRejectsCultureDependentNumbersAndTimezoneConversion() {
        var fields = new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.DurationMinutes, ProjectDataField.Start };
        foreach (var values in new[] { new[] { "1", "Task", "1,5", "" }, new[] { "1", "Task", "60", "2026-10-05T08:00:00Z" } }) {
            var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] { values });
            var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, fields.Select(f => new ProjectDataColumn(f, f.ToString())));
            Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped }));
        }
    }
}
