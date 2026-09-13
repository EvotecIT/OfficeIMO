using OfficeIMO;
using OfficeIMO.Project;
using System.Text.Json;

/// <summary>Every adopted format pair exercises explicit assessment, commit, detection, and stable core identities.</summary>
internal static class ProjectConversionProof {
    internal static int Run(string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var formats = Enum.GetValues<ProjectFileFormat>().Where(f => f != ProjectFileFormat.Automatic).ToArray();
        static string Extension(ProjectFileFormat format) => format == ProjectFileFormat.Xml ? ".xml" : format == ProjectFileFormat.Mpx4 ? ".mpx" : format.ToString().StartsWith("Mpt", StringComparison.Ordinal) ? ".mpt" : ".mpp";
        var results = new List<object>(); int failures = 0;
        foreach (var source in formats) {
            string sourcePath = Path.Combine(output, "source-" + source + Extension(source));
            using (var authored = MpxLifecycleProof.Create()) {
                // A one-owner resource calendar is supported by every adopted native generation.
                authored.Resources[0].Calendar = authored.Calendars.Add("Engineer", authored.Calendar);
                var options = new ProjectSaveOptions { Format = source, LossPolicy = OfficeConversionLossPolicy.Allow };
                authored.AssessSave(options).ThrowIfErrors(); authored.Save(sourcePath, options);
            }
            foreach (var target in formats) {
                string path = Path.Combine(output, source + "-to-" + target + Extension(target));
                try {
                    using var document = ProjectDocument.Load(sourcePath);
                    if (document.Settings.CurrencyCode == null) document.Settings.CurrencyCode = "USD"; // The synthetic input is explicitly USD.
                    var names = document.AllTasks.ToDictionary(t => t.Uid, t => t.Name);
                    var parents = document.AllTasks.ToDictionary(t => t.Uid, t => t.Parent?.Uid);
                    var assignments = document.Assignments.ToDictionary(a => (a.Task!.Uid, a.Resource!.Uid), a => (a.Work, a.Cost));
                    var options = new ProjectSaveOptions { Format = target, LossPolicy = OfficeConversionLossPolicy.Allow };
                    var report = document.AssessSave(path, options); report.ThrowIfErrors(); document.Save(path, options);
                    using var reopened = ProjectDocument.Load(path);
                    var actualTasks = reopened.AllTasks.ToArray();
                    var generated = actualTasks.Where(t => !names.ContainsKey(t.Uid)).ToArray();
                    if (actualTasks.Length - generated.Length != names.Count || generated.Any(t => reopened.NativeInfo == null || t.Uid != 0 || !t.IsSummary || t.Parent != null || t.SourceOutlineLevel != 0))
                        throw new InvalidDataException("Task count changed outside the qualified native project-summary row.");
                    foreach (var task in actualTasks.Where(t => names.ContainsKey(t.Uid))) if (names[task.Uid] != task.Name || parents[task.Uid] != task.Parent?.Uid) throw new InvalidDataException("Task identity or hierarchy changed.");
                    if (reopened.Assignments.Count != assignments.Count) throw new InvalidDataException("Assignment count changed.");
                    foreach (var assignment in reopened.Assignments) if (!assignments[(assignment.Task!.Uid, assignment.Resource!.Uid)].Equals((assignment.Work, assignment.Cost))) throw new InvalidDataException("Assignment identity, work, or cost changed.");
                    if (target != ProjectFileFormat.Xml && target != ProjectFileFormat.Mpx4 && reopened.NativeInfo!.Profile.Format(reopened.NativeInfo.IsTemplate) != target)
                        throw new InvalidDataException("The native target generation/template marker is incorrect.");
                    results.Add(new { source = source.ToString(), target = target.ToString(), passed = true, generatedProjectSummary = generated.Length == 1, losses = report.Diagnostics.Where(d => d.RepresentsLoss).Select(d => new { d.Code, d.Location }) });
                } catch (Exception error) { failures++; results.Add(new { source = source.ToString(), target = target.ToString(), passed = false, error = error.Message }); Console.Error.WriteLine(source + " -> " + target + ": " + error.Message); }
            }
        }
        File.WriteAllText(Path.Combine(output, "matrix.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("Format pairs=" + results.Count + " failures=" + failures); return failures == 0 ? 0 : 1;
    }
}
