using OfficeIMO.Drawing;
using OfficeIMO.Project;
using OfficeIMO.Workflows;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a calculated delivery plan, a progress Gantt, and an editable Word report.</summary>
internal static class ProjectDeliveryGantt {
    internal static void Create(string folder) {
        using var project = ProjectDocument.Create();
        project.Name = "Service launch / delivery plan";
        project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);

        var scope = project.Tasks.Add("Agree scope and acceptance");
        var design = project.Tasks.Add("Design the service");
        var build = project.Tasks.Add("Build and integrate");
        var verify = project.Tasks.Add("Verify the release");
        var launch = project.Tasks.Add("Launch milestone");
        scope.Duration = ProjectDuration.WorkingDays(2);
        design.Duration = ProjectDuration.WorkingDays(2);
        build.Duration = ProjectDuration.WorkingDays(4);
        verify.Duration = ProjectDuration.WorkingDays(2);
        launch.Duration = ProjectDuration.WorkingHours(0);
        project.Dependencies.Add(scope, design);
        project.Dependencies.Add(design, build);
        project.Dependencies.Add(build, verify);
        project.Dependencies.Add(verify, launch);
        scope.ActualDuration = ProjectDuration.WorkingHours(16);
        scope.RemainingDuration = ProjectDuration.WorkingHours(0);
        design.ActualDuration = ProjectDuration.WorkingHours(8);
        design.RemainingDuration = ProjectDuration.WorkingHours(8);

        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions {
            Kind = ProjectViewKind.Gantt, Timescale = ProjectViewTimescale.Day,
            PageWidth = 1200, PageHeight = 675,
            StatusDate = new DateTime(2026, 10, 8),
            Columns = new[] { ProjectViewColumn.Uid, ProjectViewColumn.Name }
        });
        project.ApplySchedule(schedule);
        project.Save(Path.Combine(folder, "example.xml"));

        // The checkout runner copies these bundled, openly licensed font faces.
        string fonts = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts");
        var faces = new OfficeFontFaceCollection()
            .Add("Arial", File.ReadAllBytes(Path.Combine(fonts, "Carlito-Regular.ttf")))
            .Add("Arial", File.ReadAllBytes(Path.Combine(fonts, "Carlito-Bold.ttf")), OfficeFontStyle.Bold);
        var typography = new OfficeRenderingProfile("delivery-report", faces, OfficeManagedTextShapingProvider.Instance);
        var images = new ProjectImageExportOptions().UseRenderingProfile(typography).UseQuality(OfficeImageExportQuality.Print);
        File.WriteAllBytes(Path.Combine(folder, "preview.png"), ProjectReportWorkflow.ExportImages(view, options: images).Single().Bytes);
        File.WriteAllText(Path.Combine(folder, "preview.svg"), ProjectReportWorkflow.ToSvg(view, typography).Single());
        File.WriteAllBytes(Path.Combine(folder, "preview.pdf"), ProjectReportWorkflow.ToPdf(view, typography: typography));
        using var word = ProjectReportWorkflow.CreateWord(view, new ProjectOfficeReportOptions { Images = images });
        word.Save(Path.Combine(folder, "report.docx"));
    }
}
