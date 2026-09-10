using OfficeIMO.Project;

if (args.Length == 3 && args[0] == "schema") return SchemaValidation.Run(args[1], args[2]);
if (args.Length == 3 && args[0] == "edit-cases") return EditCases.Run(args[1], args[2]);
if (args.Length == 3 && args[0] == "cancellation") return CancellationProof.Run(args[1], args[2]);
if (args.Length == 3 && args[0] == "native-probe") return NativeProbe.Run(args[1], args[2]);
if (args.Length == 4 && (args[0] == "scale-create" || args[0] == "scale-read-edit-save"))
    return ScaleWorkload.Run(args[0], args[1], int.Parse(args[2], System.Globalization.CultureInfo.InvariantCulture), args[3]);
if (args.Length == 3 && args[0] == "corpus") {
    if (Directory.Exists(args[2])) throw new IOException("Choose a new output directory.");
    Directory.CreateDirectory(args[2]);
    foreach (string file in Directory.GetFiles(args[1], "*.xml")) {
        using var document = ProjectDocument.Load(file);
        document.Validate().ThrowIfErrors();
        using var unchanged = new MemoryStream();
        document.Save(unchanged);
        if (!File.ReadAllBytes(file).SequenceEqual(unchanged.ToArray())) throw new InvalidDataException("Unchanged fixture differs: " + file);
        var build = document.AllTasks.FirstOrDefault(t => t.Name == "Build");
        if (build != null) { build.Name = "Build revised"; build.Notes = "OfficeIMO edit: café / Łódź / 日本語"; }
        else document.Title = "OfficeIMO empty edit";
        document.Save(Path.Combine(args[2], Path.GetFileName(file)));
        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { fixture = Path.GetFileName(file),
            tasks = document.AllTasks.Count(), calendars = document.Calendars.Count, assignments = document.Assignments.Count,
            buildCost = build?.Cost, fixedCost = build?.FixedCost, costPerUse = document.Resources.FirstOrDefault(r => r.Name == "Engineer")?.CostPerUse }));
    }
    return 0;
}

if (args.Length != 2) {
    Console.Error.WriteLine("Usage: OfficeIMO.Project.Verification <source.xml> <new-output-directory>");
    return 2;
}
string input = Path.GetFullPath(args[0]);
string output = Path.GetFullPath(args[1]);
if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
Directory.CreateDirectory(output);
using (var imported = ProjectDocument.Load(input)) {
    imported.Validate().ThrowIfErrors();
    string unchanged = Path.Combine(output, "unchanged.xml");
    imported.Save(unchanged);
    if (!File.ReadAllBytes(input).SequenceEqual(File.ReadAllBytes(unchanged))) throw new InvalidDataException("Unchanged XML bytes were not preserved.");
    var build = imported.AllTasks.Single(t => t.Name == "Build");
    if (build.Cost != 5000m || build.Work?.Minutes != 2400m) throw new InvalidDataException("Imported cost/work differs from the Project object-model oracle.");
    build.Name = "Build revised";
    build.Notes = "OfficeIMO edit: café / Łódź / 日本語";
    imported.Save(Path.Combine(output, "edited.xml"));
    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new {
        imported.SourceSaveVersion, Tasks = imported.AllTasks.Count(), Resources = imported.Resources.Count,
        Assignments = imported.Assignments.Count, Calendars = imported.Calendars.Count,
        BuildCost = build.Cost, WorkMinutes = build.Work?.Minutes, PreservedDiagnostics = imported.ReadDiagnostics.Count
    }));
}
using (var authored = ProjectDocument.Create(Path.Combine(output, "authored.xml"))) {
    authored.Name = "OfficeIMO authored";
    authored.Title = "OfficeIMO authored";
    authored.Author = "OfficeIMO";
    authored.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
    authored.Settings.CurrencyCode = "PLN";
    authored.Settings.CurrencyDigits = 2;
    authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
    var engineer = authored.Resources.AddWork("Engineer");
    engineer.StandardRate = 125;
    var summary = authored.Tasks.AddSummary("Delivery");
    var design = summary.Children.Add("Design");
    design.IsManual = false;
    design.Duration = ProjectDuration.WorkingDays(3);
    var build = summary.Children.Add("Build");
    build.IsManual = false;
    build.Duration = ProjectDuration.WorkingDays(5);
    authored.Dependencies.Add(design, build);
    var assignment = authored.Assignments.Add(build, engineer);
    assignment.Work = ProjectWork.Hours(40);
    assignment.Cost = 5000;
    build.Cost = 5000;
    build.Notes = "OfficeIMO authored: café / Łódź / 日本語";
    var baseline = build.Baselines.Add(); baseline.Number = 0; baseline.Work = ProjectWork.Hours(40); baseline.Cost = 5000;
    authored.Save();
    Console.WriteLine("Authored XML: " + Path.Combine(output, "authored.xml"));
}
return 0;
