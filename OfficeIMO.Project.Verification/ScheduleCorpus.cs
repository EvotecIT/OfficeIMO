using System.Text.Json;
using System.Xml.Linq;
using System.Globalization;
using OfficeIMO.Project;

internal static class ScheduleCorpus {
    internal static int Run(string source, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var comparisons = new List<object>(); var differences = new List<object>();
        foreach (string name in new[] { "delivery", "relationships", "calendars", "constraints", "backward", "work-cost", "rich-fields", "working-weeks" }.Where(n => File.Exists(Path.Combine(source, n + ".xml")))) {
          foreach (var format in new[] { ".xml", ".mpp" }.Where(ext => File.Exists(Path.Combine(source, name + ext)))) {
            using var document = ProjectDocument.Load(Path.Combine(source, name + format));
            using var expectedDocument = ProjectDocument.Load(Path.Combine(source, name + ".xml"));
            var reference = XDocument.Load(Path.Combine(source, name + ".xml"));
            var ns = reference.Root!.Name.Namespace;
            var referenceTasks = reference.Root.Element(ns + "Tasks")!.Elements(ns + "Task").ToDictionary(t => (int)t.Element(ns + "UID")!);
            long revision = document.Revision;
            var result = document.CalculateSchedule();
            foreach (var diagnostic in result.Report.Diagnostics.Where(d => d.Severity == ProjectDiagnosticSeverity.Error))
                differences.Add(new { fixture = name, code = diagnostic.Code, location = diagnostic.Location, message = diagnostic.Message });
            foreach (var actual in result.Tasks) {
                var expected = expectedDocument.AllTasks.Single(t => t.Uid == actual.TaskUid);
                if (expected.Start != actual.Start || expected.Finish != actual.Finish)
                    differences.Add(new { fixture = name, task = expected.Name, expectedStart = expected.Start, expectedFinish = expected.Finish, actual.Start, actual.Finish });
                if (expected.Duration is ProjectDuration duration && (duration.Value != actual.Duration.Value || duration.Unit != actual.Duration.Unit || duration.IsElapsed != actual.Duration.IsElapsed))
                    differences.Add(new { fixture = name, format, task = expected.Name, field = "Duration", expected = duration, actual = actual.Duration });
                var raw = referenceTasks[actual.TaskUid];
                foreach (var property in new[] { "EarlyStart", "EarlyFinish", "LateStart", "LateFinish" }) {
                    if (raw.Element(ns + property) is not XElement field) continue;
                    var expectedDate = DateTime.Parse(field.Value, CultureInfo.InvariantCulture);
                    var actualDate = (DateTime)typeof(ProjectTaskSchedule).GetProperty(property)!.GetValue(actual)!;
                    if (expectedDate != actualDate) differences.Add(new { fixture = name, task = expected.Name, field = property, expected = expectedDate, actual = actualDate });
                }
                foreach (var pair in new[] { ("TotalSlack", actual.TotalSlackMinutes), ("FreeSlack", actual.FreeSlackMinutes) })
                    if (raw.Element(ns + pair.Item1) is XElement field && decimal.Parse(field.Value, CultureInfo.InvariantCulture) / 10m != pair.Item2)
                        differences.Add(new { fixture = name, task = expected.Name, field = pair.Item1, expected = decimal.Parse(field.Value, CultureInfo.InvariantCulture) / 10m, actual = pair.Item2 });
                if (expected.IsCritical.HasValue && expected.IsCritical != actual.IsCritical)
                    differences.Add(new { fixture = name, task = expected.Name, field = "Critical", expected = expected.IsCritical, actual = actual.IsCritical });
            }
            if (document.Revision != revision) throw new InvalidOperationException("Calculation mutated the document.");
            comparisons.Add(new { fixture = name, format, revision, results = result.Tasks, diagnostics = result.Report.Diagnostics });
          }
        }
        File.WriteAllText(Path.Combine(output, "schedule-comparison.json"), JsonSerializer.Serialize(new { comparisons, differences }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(JsonSerializer.Serialize(new { fixtures = comparisons.Count, differences = differences.Count, details = differences }));
        return differences.Count == 0 ? 0 : 1;
    }
}
