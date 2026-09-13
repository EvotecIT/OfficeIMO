using OfficeIMO.Drawing;
using OfficeIMO.Project;
using OfficeIMO.Workflows;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a release timeline and a PowerPoint report with charts and editable data.</summary>
internal static class ProjectMilestoneTimeline {
    internal static void Create(string folder) {
        using var project = ProjectDocument.Create();
        project.Name = "Autumn release / milestone timeline";
        project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        string[] names = { "Prepare the pilot", "Pilot ready", "Run the pilot", "Pilot accepted", "Roll out the service", "Service handover" };
        var tasks = names.Select((name, index) => {
            var task = project.Tasks.Add(name);
            task.Duration = ProjectDuration.WorkingDays(index % 2 == 1 ? 0 : 3);
            return task;
        }).ToArray();
        for (int index = 1; index < tasks.Length; index++)
            project.Dependencies.Add(tasks[index - 1], tasks[index]);

        var schedule = project.CalculateSchedule();
        schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions {
            Kind = ProjectViewKind.Timeline, Timescale = ProjectViewTimescale.Day,
            PageWidth = 1200, PageHeight = 675,
            Columns = new[] { ProjectViewColumn.Name }
        });
        project.ApplySchedule(schedule);
        project.Save(Path.Combine(folder, "example.xml"));

        string fonts = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts");
        var faces = new OfficeFontFaceCollection()
            .Add("Arial", File.ReadAllBytes(Path.Combine(fonts, "Carlito-Regular.ttf")))
            .Add("Arial", File.ReadAllBytes(Path.Combine(fonts, "Carlito-Bold.ttf")), OfficeFontStyle.Bold);
        var typography = new OfficeRenderingProfile("timeline-report", faces, OfficeManagedTextShapingProvider.Instance);
        var images = new ProjectImageExportOptions().UseRenderingProfile(typography).UseQuality(OfficeImageExportQuality.Print);
        File.WriteAllBytes(Path.Combine(folder, "preview.png"), ProjectReportWorkflow.ExportImages(view, options: images).Single().Bytes);
        File.WriteAllText(Path.Combine(folder, "preview.svg"), ProjectReportWorkflow.ToSvg(view, typography).Single());
        File.WriteAllBytes(Path.Combine(folder, "preview.pdf"), ProjectReportWorkflow.ToPdf(view, typography: typography));
        using var slides = ProjectReportWorkflow.CreatePowerPoint(view, new ProjectOfficeReportOptions { Images = images });
        slides.Save(Path.Combine(folder, "report.pptx"));
        using var workbook = ProjectReportWorkflow.CreateExcel(view);
        workbook.Save(Path.Combine(folder, "report.xlsx"));
    }
}
