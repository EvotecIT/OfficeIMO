using System.Text;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class ProjectReportWorkflowTests {
    [Fact]
    public void EditablePresentationPaginatesLongLabelsWithoutLosingRows() {
        using var project = Create(40);
        project.Tasks[0].Name = "Long task label with implementation, verification and documented handover for the shared service";
        var view = project.CreateView(project.CalculateSchedule(), new ProjectViewOptions {
            Kind = ProjectViewKind.Table, Columns = new[] { ProjectViewColumn.Uid, ProjectViewColumn.Name }, PageHeight = 400
        });
        using var report = ProjectReportWorkflow.CreatePowerPoint(view);
        using var copy = PowerPointPresentation.Load(new MemoryStream(report.ToBytes()));
        var tables = copy.Slides.SelectMany(s => s.Shapes.OfType<PowerPointTable>()).ToArray();
        Assert.True(tables.Length > 3);
        var names = tables.SelectMany(t => t.RowItems.Skip(1)).Select(r => r.Cells[1].Text).ToArray();
        // Main report rows and supplemental status rows both retain every full name.
        foreach (var task in project.Tasks) Assert.Equal(2, names.Count(n => n == task.Name));
        Assert.All(tables, table => Assert.True(table.TopPoints + table.HeightPoints <= view.PageHeight - view.PageMargin));
        Assert.Empty(copy.ValidateDocument());
    }

    [Fact]
    public void ExcelIntegralDecimalScaleIsNormalizedBeforeIdentityImport() {
        using var workbook = ExcelDocument.Create(); var sheet = workbook.AddWorksheet("Tasks");
        sheet.CellValue(1, 1, "Uid"); sheet.CellValue(1, 2, "Name");
        sheet.CellValue(2, 1, 1.0m); sheet.CellValue(2, 2, "Task");
        Assert.True(sheet.TryGetCellValueSnapshot(2, 1, out var raw)); Assert.Equal("1.0", raw!.RawValue);
        var table = ProjectDataWorkflow.ReadExcel(sheet, 1, 2);
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, new[] { ProjectDataField.Uid, ProjectDataField.Name }.Select(f => new ProjectDataColumn(f, f.ToString())));
        using var project = ProjectDocument.ImportTables(new[] { mapped }).Document;
        Assert.Equal(1, project.Tasks[0].Uid);
    }

    [Theory]
    [InlineData(1.0, "1")]
    [InlineData(1.25, "1.25")]
    [InlineData(1e-28, "0.0000000000000000000000000001")]
    public void NativeExcelNumbersNormalizeWithoutLosingPrecision(double number, string expected) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Tasks");
        sheet.CellValue(1, 1, "Uid"); sheet.CellValue(1, 2, "Name");
        sheet.CellValue(2, 1, number); sheet.CellValue(2, 2, "Task");
        var table = ProjectDataWorkflow.ReadExcel(sheet, 1, 2);
        Assert.Equal(expected, table.Rows[0][0]);
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, new[] { ProjectDataField.Uid, ProjectDataField.Name }.Select(f => new ProjectDataColumn(f, f.ToString())));
        if (number == 1) { using var project = ProjectDocument.ImportTables(new[] { mapped }).Document; Assert.Equal(1, project.Tasks[0].Uid); }
        else Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped }));
    }

    [Theory]
    [InlineData(1e-29)]
    [InlineData(-1e-29)]
    public void NativeExcelNumbersRejectDecimalUnderflow(double number) {
        using var workbook = ExcelDocument.Create(); var sheet = workbook.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Cost"); sheet.CellValue(2, 1, number);
        Assert.Throws<InvalidDataException>(() => ProjectDataWorkflow.ReadExcel(sheet, 1, 1));
    }

    [Fact]
    public void ReportsProduceReadablePdfAndEditableOfficePackages() {
        using var project = Create(18);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage });
        byte[] pdf = ProjectReportWorkflow.ToPdf(view);
        Assert.Equal(view.Render().Count, PdfDocument.Load(pdf).Read().PageCount);
        string html = ProjectReportWorkflow.ToHtml(view);
        Assert.Contains("<svg", html); Assert.Contains("<table", html); Assert.Contains("&lt;unsafe&gt;", html);
        using var word = ProjectReportWorkflow.CreateWord(view);
        Assert.Empty(word.ValidateDocument());
        using var wordCopy = WordDocument.Load(new MemoryStream(word.ToBytes()));
        Assert.True(wordCopy.Tables.Count > 1);
        Assert.Contains(wordCopy.Tables.SelectMany(t => t.Rows).SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs), p => p.Text.Contains("Task 0"));
        using var presentation = ProjectReportWorkflow.CreatePowerPoint(view);
        Assert.Empty(presentation.ValidateDocument());
        using var presentationCopy = PowerPointPresentation.Load(new MemoryStream(presentation.ToBytes()));
        Assert.True(presentationCopy.Slides.Count > 1);
        Assert.All(presentationCopy.Slides, slide => Assert.Single(slide.Shapes.OfType<PowerPointTable>()));
        using var workbook = ProjectReportWorkflow.CreateExcel(view);
        using var workbookCopy = ExcelDocument.Load(new MemoryStream(workbook.ToBytes()));
        Assert.Empty(workbookCopy.ValidateDocument());
        var usage = workbookCopy.Sheets.Single(s => s.Name == "Work hours");
        Assert.True(usage.TryGetCellValueSnapshot(2, 3, out var value));
        Assert.Equal(ExcelCellValueKind.Number, value!.Kind); Assert.Equal(8m, decimal.Parse(value.RawValue, System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void CsvAndExcelTransportKeepLiteralTextAndMappedIdentities() {
        using var project = Create(1);
        project.Tasks[0].Name = "=1+1, \"quoted\"\nZażółć";
        var export = project.ExportTables(allowLossyProjection: true);
        var taskTable = export.Tables.Single(t => t.Kind == ProjectDataKind.Tasks);
        var csv = ProjectDataWorkflow.CreateCsv(taskTable.Table);
        using var csvStream = new MemoryStream(); csv.Save(csvStream); csvStream.Position = 0;
        var csvRead = ProjectDataWorkflow.ReadCsv(CsvDocument.Load(csvStream));
        Assert.Equal(project.Tasks[0].Name, csvRead.Rows[0][1]);
        using var workbook = ProjectDataWorkflow.CreateExcel(export);
        using var reloaded = ExcelDocument.Load(new MemoryStream(workbook.ToBytes()));
        Assert.Empty(reloaded.ValidateDocument());
        var mapped = export.Tables.Select(table => new ProjectMappedTable(table.Kind,
            ProjectDataWorkflow.ReadExcel(reloaded.Sheets.Single(s => s.Name == table.Kind.ToString()), table.Table.Rows.Count, table.Table.Headers.Count), table.Columns)).ToArray();
        var imported = ProjectDocument.ImportTables(mapped, new ProjectDataImportOptions { Start = project.Settings.StartDate, CalendarUid = project.Calendar!.Uid });
        using var result = imported.Document;
        Assert.Equal(project.Tasks[0].Name, result.Tasks.GetByUid(project.Tasks[0].Uid).Name);
        Assert.Empty(reloaded.Sheets.Single(s => s.Name == "Tasks").GetFormulaCells());
    }

    [Fact]
    public void RasterAndCancellationUseTheSharedRendererContract() {
        using var project = Create(1);
        var view = project.CreateView(project.CalculateSchedule());
        byte[] png = ProjectReportWorkflow.ToPng(view).Single();
        Assert.Equal(new byte[] { 137, 80, 78, 71 }, png.Take(4));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => ProjectReportWorkflow.ToPng(view, cancellationToken: cancel.Token));
    }

    private static ProjectDocument Create(int count) {
        var project = ProjectDocument.Create(); project.Name = "Report <unsafe>";
        project.Calendar = project.Calendars.AddStandardWorkingWeek(); project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var resource = project.Resources.AddWork("Engineer");
        for (int i = 0; i < count; i++) {
            var task = project.Tasks.Add("Task " + i); task.Duration = ProjectDuration.WorkingMinutes(480);
            project.Assignments.Add(task, resource);
        }
        return project;
    }
}
