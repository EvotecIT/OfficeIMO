using OfficeIMO.Drawing;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows.Tests;

public sealed class ProjectImageQualityTests {
    [Fact]
    public void TableOnlyViewRetainsItsContentWhenChartSelectionIsUsed() {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        project.Tasks.Add("Editable task").Duration = ProjectDuration.WorkingHours(8);
        var view = project.CreateView(project.CalculateSchedule(), options: new ProjectViewOptions { Kind = ProjectViewKind.Table });
        var options = new ProjectOfficeReportOptions { IncludeDataTables = false };
        using var word = ProjectReportWorkflow.CreateWord(view, options);
        using var slides = ProjectReportWorkflow.CreatePowerPoint(view, options);
        Assert.NotEmpty(word.Tables);
        Assert.Contains(slides.Slides.SelectMany(slide => slide.Shapes), shape => shape is OfficeIMO.PowerPoint.PowerPointTable);
    }

    [Fact]
    public void ImageExportEnforcesCancellationAndPixelBudget() {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        project.Tasks.Add("Task").Duration = ProjectDuration.WorkingHours(8);
        var view = project.CreateView(project.CalculateSchedule(), options: new ProjectViewOptions { Kind = ProjectViewKind.Table });
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => ProjectReportWorkflow.ExportImages(view, cancellationToken: cancellation.Token));
        Assert.Throws<OfficeImageExportLimitException>(() => ProjectReportWorkflow.ExportImages(view, options: new ProjectImageExportOptions { MaximumRasterPixels = 1 }));
    }

    [Fact]
    public void ReportImagesUsePointDensityAndOfficeReportsContainChartsAndEditableTables() {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        project.Tasks.Add("Delivery").Duration = ProjectDuration.WorkingHours(8);
        var view = project.CreateView(project.CalculateSchedule());
        var options = new ProjectImageExportOptions().UseQuality(OfficeImageExportQuality.Preview);
        var image = Assert.Single(ProjectReportWorkflow.ExportImages(view, options: options));
        Assert.Equal((int)Math.Ceiling(view.PageWidth * 96 / 72), image.Width);
        Assert.Equal(1, options.Scale);
        using var word = ProjectReportWorkflow.CreateWord(view, new ProjectOfficeReportOptions { Images = options });
        Assert.NotEmpty(word.Images); Assert.NotEmpty(word.Tables); Assert.Empty(word.ValidateDocument());
        using var slides = ProjectReportWorkflow.CreatePowerPoint(view, new ProjectOfficeReportOptions { Images = options });
        Assert.Contains(slides.Slides.SelectMany(slide => slide.Shapes), shape => shape is OfficeIMO.PowerPoint.PowerPointPicture);
        Assert.Contains(slides.Slides.SelectMany(slide => slide.Shapes), shape => shape is OfficeIMO.PowerPoint.PowerPointTable);
        Assert.Empty(slides.ValidateDocument());
        using var tables = ProjectReportWorkflow.CreateWord(view, new ProjectOfficeReportOptions { IncludeCharts = false });
        Assert.Empty(tables.Images); Assert.NotEmpty(tables.Tables);
    }
}
