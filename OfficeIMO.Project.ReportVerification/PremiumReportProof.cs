using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Project;
using OfficeIMO.Workflows;

internal static class PremiumReportProof {
    internal static int Run(string output, OfficeRenderingProfile typography) {
        using var project = ProjectDocument.Create(); project.Name = "Platform release · delivery plan";
        project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        string[] names = { "Discovery & scope", "API design", "Interface design", "Data migration plan", "Build & integrate", "Migration rehearsal", "Acceptance & launch readiness", "Release" };
        var tasks = names.Select((name, index) => { var task = project.Tasks.Add(name); task.Duration = ProjectDuration.WorkingDays(index == 7 ? 0 : index == 4 ? 4 : 2); return task; }).ToArray();
        foreach (var edge in new[] { (0, 1), (0, 2), (0, 3), (1, 4), (2, 4), (3, 5), (4, 6), (5, 6), (6, 7) }) project.Dependencies.Add(tasks[edge.Item1], tasks[edge.Item2]);
        tasks[0].ActualDuration = ProjectDuration.WorkingHours(16); tasks[0].RemainingDuration = ProjectDuration.WorkingHours(0);
        tasks[1].ActualDuration = ProjectDuration.WorkingHours(8); tasks[1].RemainingDuration = ProjectDuration.WorkingHours(8);
        tasks[2].ActualDuration = ProjectDuration.WorkingHours(4); tasks[2].RemainingDuration = ProjectDuration.WorkingHours(12);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var evidence = new List<object>();
        foreach (var kind in new[] { ProjectViewKind.Gantt, ProjectViewKind.Network, ProjectViewKind.Timeline }) {
            var view = project.CreateView(schedule, new ProjectViewOptions { Kind = kind, PageWidth = kind == ProjectViewKind.Network ? 1600 : 1200,
                PageHeight = 900, Timescale = ProjectViewTimescale.Day, StatusDate = new DateTime(2026, 10, 8),
                Columns = new[] { ProjectViewColumn.Uid, ProjectViewColumn.Name }, MaxPages = 30 });
            var options = new ProjectImageExportOptions().UseRenderingProfile(typography);
            var images = ProjectReportWorkflow.ExportImages(view, options: options);
            var svgs = ProjectReportWorkflow.ToSvg(view, typography);
            for (int p = 0; p < images.Count; p++) {
                File.WriteAllBytes(Path.Combine(output, kind + "-" + (p + 1) + ".png"), images[p].Bytes);
                File.WriteAllText(Path.Combine(output, kind + "-" + (p + 1) + ".svg"), svgs[p]);
            }
            File.WriteAllText(Path.Combine(output, kind + ".html"), ProjectReportWorkflow.ToHtml(view, typography));
            File.WriteAllBytes(Path.Combine(output, kind + ".pdf"), ProjectReportWorkflow.ToPdf(view, typography: typography));
            evidence.Add(new { kind = kind.ToString(), pages = images.Count, width = images[0].Width, height = images[0].Height, dpi = images[0].DpiX });
            if (kind == ProjectViewKind.Gantt) {
                var native = new ProjectOfficeReportOptions { Images = options };
                using var word = ProjectReportWorkflow.CreateWord(view, native); word.Save(Path.Combine(output, "report.docx"));
                using var presentation = ProjectReportWorkflow.CreatePowerPoint(view, native); presentation.Save(Path.Combine(output, "report.pptx"));
                using var excel = ProjectReportWorkflow.CreateExcel(view); excel.Save(Path.Combine(output, "report.xlsx"));
                var errors = word.ValidateDocument().Select(error => "Word: " + error.Description)
                    .Concat(presentation.ValidateDocument().Select(error => "PowerPoint: " + error.Description))
                    .Concat(excel.ValidateDocument().Select(error => "Excel: " + error.Description)).ToArray();
                if (errors.Length != 0) throw new InvalidDataException(string.Join("\n", errors));
            }
        }
        File.WriteAllText(Path.Combine(output, "evidence.json"), JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(JsonSerializer.Serialize(evidence));
        return 0;
    }
}
