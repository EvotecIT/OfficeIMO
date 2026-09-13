using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

if (args.Length < 2 || args.Length > 4 || args.Length >= 3 && args[2] != "native" && args[2] != "premium") {
    Console.Error.WriteLine("Usage: OfficeIMO.Project.ReportVerification <new-output-directory> <font.ttf> [native|premium] [bold-font.ttf]"); return 2;
}
bool nativeOnly = args.Length >= 3 && args[2] == "native";
string output = Path.GetFullPath(args[0]);
if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
var fonts = new OfficeFontFaceCollection().Add("Arial", File.ReadAllBytes(args[1]));
if (args.Length == 4) fonts.Add("Arial", File.ReadAllBytes(args[3]), OfficeFontStyle.Bold);
var typography = new OfficeRenderingProfile("report-proof", fonts, OfficeManagedTextShapingProvider.Instance);
Directory.CreateDirectory(output);
if (args.Length >= 3 && args[2] == "premium") return PremiumReportProof.Run(output, typography);
using var project = ProjectDocument.Create(); project.Name = "Delivery plan · Łódź";
project.Calendar = project.Calendars.AddStandardWorkingWeek(); project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
var engineer = project.Resources.AddWork("Engineering"); engineer.StandardRate = 100;
var analyst = project.Resources.AddWork("Analysis"); analyst.StandardRate = 120;
var summary = project.Tasks.AddSummary("Release preparation");
ProjectTask? previous = null;
string[] names = { "Define requirements", "Architecture & interfaces", "Zażółć — implementation", "Integration tests", "Accessibility review", "Documentation", "Release approval", "Launch" };
for (int i = 0; i < names.Length; i++) {
    var task = summary.Children.Add(names[i]); task.Duration = ProjectDuration.WorkingHours(i == 7 ? 0 : 8 + i % 3 * 4);
    if (previous != null) project.Dependencies.Add(previous, task);
    if (i != 7) project.Assignments.Add(task, i % 2 == 0 ? engineer : analyst);
    var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Start = project.Settings.StartDate!.Value.AddDays(i); baseline.Finish = baseline.Start.Value.AddHours(8);
    previous = task;
}
var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
var evidence = new List<object>();
foreach (var kind in nativeOnly ? Array.Empty<ProjectViewKind>() : Enum.GetValues<ProjectViewKind>()) {
    var view = project.CreateView(schedule, new ProjectViewOptions { Kind = kind, Timescale = ProjectViewTimescale.Day, BaselineNumber = 0 });
    string stem = Path.Combine(output, kind.ToString());
    File.WriteAllText(stem + ".html", ProjectReportWorkflow.ToHtml(view, typography));
    var svgs = ProjectReportWorkflow.ToSvg(view, typography); var pngs = ProjectReportWorkflow.ToPng(view, typography: typography);
    for (int i = 0; i < svgs.Count; i++) { File.WriteAllText(stem + "-" + (i + 1) + ".svg", svgs[i]); File.WriteAllBytes(stem + "-" + (i + 1) + ".png", pngs[i]); }
    byte[] pdf = ProjectReportWorkflow.ToPdf(view, typography: typography); File.WriteAllBytes(stem + ".pdf", pdf);
    int pdfPages = PdfDocument.Load(pdf).Read().PageCount;
    if (pdfPages != svgs.Count) throw new InvalidDataException("PDF page count differs for " + kind);
    evidence.Add(new { kind = kind.ToString(), rows = view.Rows.Count, buckets = view.Buckets.Count, pages = svgs.Count, pdfPages, pngPages = pngs.Count });
}
var usageView = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage, Timescale = ProjectViewTimescale.Day, BaselineNumber = 0 });
using (var word = ProjectReportWorkflow.CreateWord(usageView)) {
    word.Save(Path.Combine(output, "report.docx"));
    using var copy = WordDocument.Load(Path.Combine(output, "report.docx"));
    if (copy.ValidateDocument().Count != 0) throw new InvalidDataException("Word report failed package validation.");
}
using (var presentation = ProjectReportWorkflow.CreatePowerPoint(usageView)) {
    presentation.Save(Path.Combine(output, "report.pptx"));
    using var copy = PowerPointPresentation.Load(Path.Combine(output, "report.pptx"));
    if (copy.ValidateDocument().Count != 0) throw new InvalidDataException("PowerPoint report failed package validation.");
}
using (var workbook = ProjectReportWorkflow.CreateExcel(usageView)) {
    workbook.Save(Path.Combine(output, "report.xlsx"));
    using var copy = ExcelDocument.Load(Path.Combine(output, "report.xlsx"));
    if (copy.ValidateDocument().Count != 0) throw new InvalidDataException("Excel report failed package validation.");
}
if (!nativeOnly) {
    var transfer = project.ExportTables(allowLossyProjection: true);
    using (var workbook = ProjectDataWorkflow.CreateExcel(transfer)) workbook.Save(Path.Combine(output, "project-data.xlsx"));
    foreach (var table in transfer.Tables) ProjectDataWorkflow.CreateCsv(table.Table).Save(Path.Combine(output, table.Kind + ".csv"));
    var empty = project.CreateView(schedule, new ProjectViewOptions { TaskUids = Array.Empty<int>() });
    File.WriteAllText(Path.Combine(output, "empty.html"), ProjectReportWorkflow.ToHtml(empty, typography));
    File.WriteAllBytes(Path.Combine(output, "empty.png"), ProjectReportWorkflow.ToPng(empty, typography: typography).Single());
    var months = project.CreateView(schedule, new ProjectViewOptions { Timescale = ProjectViewTimescale.Month, Start = new DateTime(2026, 9, 15), Finish = new DateTime(2026, 12, 15), BaselineNumber = 0 });
    File.WriteAllText(Path.Combine(output, "months.html"), ProjectReportWorkflow.ToHtml(months, typography));
    var monthPages = ProjectReportWorkflow.ToPng(months, typography: typography);
    for (int i = 0; i < monthPages.Count; i++) File.WriteAllBytes(Path.Combine(output, "months-" + (i + 1) + ".png"), monthPages[i]);
}
using var large = ProjectDocument.Create(); large.Name = "Large delivery plan · Łódź";
large.Calendar = large.Calendars.AddStandardWorkingWeek(); large.Settings.StartDate = project.Settings.StartDate;
ProjectTask? priorLarge = null;
for (int i = 0; i < 75; i++) {
    var task = large.Tasks.Add(i % 10 == 0 ? "Long task label — design, implementation, verification and documented handover for the shared service " + i : "Delivery item " + i);
    task.Duration = ProjectDuration.WorkingHours(4);
    if (priorLarge != null) large.Dependencies.Add(priorLarge, task);
    priorLarge = task;
}
var largeSchedule = large.CalculateSchedule(); largeSchedule.Report.ThrowIfErrors();
foreach (var kind in nativeOnly ? Array.Empty<ProjectViewKind>() : new[] { ProjectViewKind.Gantt, ProjectViewKind.Network, ProjectViewKind.Table }) {
    var view = large.CreateView(largeSchedule, new ProjectViewOptions { Kind = kind });
    string stem = Path.Combine(output, "large-" + kind);
    File.WriteAllText(stem + ".html", ProjectReportWorkflow.ToHtml(view, typography));
    var images = ProjectReportWorkflow.ToPng(view, typography: typography);
    foreach (int index in new[] { 0, images.Count / 2, images.Count - 1 }.Distinct()) File.WriteAllBytes(stem + "-" + (index + 1) + ".png", images[index]);
    byte[] pdf = ProjectReportWorkflow.ToPdf(view, typography: typography); File.WriteAllBytes(stem + ".pdf", pdf);
    int pdfPages = PdfDocument.Load(pdf).Read().PageCount;
    if (pdfPages != images.Count) throw new InvalidDataException("Large report pagination differs for " + kind);
    evidence.Add(new { kind = "large-" + kind, rows = view.Rows.Count, buckets = view.Buckets.Count, pages = images.Count, pdfPages, pngPages = images.Count });
}
var largeView = large.CreateView(largeSchedule, new ProjectViewOptions { Kind = ProjectViewKind.Table, Columns = Enum.GetValues<ProjectViewColumn>() });
foreach (var kind in nativeOnly ? Array.Empty<ProjectViewKind>() : new[] { ProjectViewKind.Gantt, ProjectViewKind.Timeline, ProjectViewKind.Table }) {
    var narrow = large.CreateView(largeSchedule, new ProjectViewOptions { Kind = kind, PageWidth = 420, PageHeight = 595,
        Columns = new[] { ProjectViewColumn.Uid, ProjectViewColumn.Name }, TaskUids = large.Tasks.Take(3).Select(task => task.Uid).ToArray() });
    File.WriteAllText(Path.Combine(output, "narrow-" + kind + ".html"), ProjectReportWorkflow.ToHtml(narrow, typography));
    File.WriteAllBytes(Path.Combine(output, "narrow-" + kind + ".png"), ProjectReportWorkflow.ToPng(narrow, typography: typography).First());
}
using (var presentation = ProjectReportWorkflow.CreatePowerPoint(largeView)) {
    Directory.CreateDirectory(Path.Combine(output, "large-native"));
    presentation.Save(Path.Combine(output, "large-native", "report.pptx"));
    if (presentation.ValidateDocument().Count != 0) throw new InvalidDataException("Large PowerPoint report failed package validation.");
}
using (var word = ProjectReportWorkflow.CreateWord(largeView)) word.Save(Path.Combine(output, "large-native", "report.docx"));
using (var excel = ProjectReportWorkflow.CreateExcel(largeView)) excel.Save(Path.Combine(output, "large-native", "report.xlsx"));
File.WriteAllText(Path.Combine(output, "evidence.json"), JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(JsonSerializer.Serialize(new { output, views = evidence.Count, files = Directory.GetFiles(output).Length, bytes = Directory.GetFiles(output).Sum(f => new FileInfo(f).Length) }));
return 0;
