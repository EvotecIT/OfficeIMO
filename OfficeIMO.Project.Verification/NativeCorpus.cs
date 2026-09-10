using OfficeIMO.Project;
using System.Text.Json;

internal static class NativeCorpus {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var failures = new List<string>(); var observations = new List<object>(); int comparisons = 0;
        foreach (var path in Directory.GetFiles(input, "*.mpp")) {
            string name = Path.GetFileName(path);
            using var native = ProjectDocument.Load(path);
            using var xml = ProjectDocument.Load(Path.ChangeExtension(path, ".xml"));
            var totals = native.AnalyzeAssignments().Resources.ToDictionary(r => r.ResourceUid);
            using var saved = new MemoryStream(); native.Save(saved);
            if (!saved.ToArray().SequenceEqual(File.ReadAllBytes(path))) failures.Add(name + ": unchanged bytes differ");
            void Compare(object? actual, object? expected, string field, string? representation = null) {
                comparisons++;
                if (Equals(actual, expected)) return;
                observations.Add(new { fixture = name, field, native = actual, xml = expected, classification = representation ?? "unexplained" });
                if (representation == null) failures.Add(name + " " + field + ": " + JsonSerializer.Serialize(actual) + " != " + JsonSerializer.Serialize(expected));
            }
            void Baselines(IEnumerable<ProjectBaseline> actuals, IEnumerable<ProjectBaseline> expecteds, string owner) {
                foreach (var baseline in expecteds) {
                    var actual = actuals.SingleOrDefault(b => b.Number == baseline.Number);
                    if (actual == null) { failures.Add(name + " baseline missing " + owner + "/" + baseline.Number); continue; }
                    string location = owner + "/Baseline[" + baseline.Number + "]";
                    Compare(actual.Start, baseline.Start, location + "/Start"); Compare(actual.Finish, baseline.Finish, location + "/Finish");
                    Compare(actual.Work, baseline.Work, location + "/Work", baseline.Work == null && actual.Work?.Minutes == 0 ? "xml-omits-explicit-zero" : null);
                    Compare(actual.Cost, baseline.Cost, location + "/Cost", baseline.Cost == null && actual.Cost == 0 ? "xml-omits-explicit-zero" : null);
                }
            }
            void Custom(IEnumerable<ProjectCustomFieldValue> actuals, IEnumerable<ProjectCustomFieldValue> expecteds, string owner) {
                foreach (var expected in expecteds.Where(v => v.Value != null && v.ValueId == null && v.ValueGuid == null)) {
                    var actual = actuals.SingleOrDefault(v => v.FieldId == expected.FieldId);
                    // Project exports enterprise flags in these synthetic local-field fixtures.
                    // Enterprise custom-field records are outside the local scalar profile.
                    Compare(actual?.Value, expected.Value, owner + "/Custom[" + expected.FieldId + "]",
                        actual == null && (expected.FieldId == "188744339" || expected.FieldId == "205521403") ? "unqualified-enterprise-field-preserved-in-native-source" : null);
                }
            }
            foreach (var expected in xml.AllTasks) {
                var task = native.AllTasks.SingleOrDefault(t => t.Uid == expected.Uid);
                if (task == null) { failures.Add(name + " task missing " + expected.Uid); continue; }
                foreach (var property in new[] { "Name", "DisplayId", "Start", "Finish", "Work", "ActualWork", "RemainingWork", "Cost", "ActualCost", "FixedCost", "RemainingCost", "PercentComplete", "PercentWorkComplete", "IsManual", "IsActive", "IsMilestone", "IsSummary", "IsCritical", "ConstraintType", "ConstraintDate", "Type", "Deadline" }) {
                    var info = typeof(ProjectTask).GetProperty(property)!;
                    Compare(info.GetValue(task), info.GetValue(expected), "task " + task.Uid + "/" + property);
                }
                Compare(task.Parent?.Uid, expected.Parent?.Uid, "task " + task.Uid + "/Parent");
                Compare(task.Calendar?.Uid, expected.Calendar?.Uid, "task " + task.Uid + "/Calendar");
                Baselines(task.Baselines, expected.Baselines, "task " + task.Uid);
                Custom(task.CustomFields, expected.CustomFields, "task " + task.Uid);
            }
            foreach (var expected in xml.Resources) {
                var resource = native.Resources.SingleOrDefault(r => r.Uid == expected.Uid);
                if (resource == null) { failures.Add(name + " resource missing " + expected.Uid); continue; }
                foreach (var property in new[] { "Name", "Type", "Work", "ActualWork", "RemainingWork", "Cost", "ActualCost", "StandardRate", "OvertimeRate", "CostPerUse", "MaxUnits" }) {
                    var info = typeof(ProjectResource).GetProperty(property)!;
                    var actualValue = info.GetValue(resource); var expectedValue = info.GetValue(expected);
                    string? classification = null;
                    if (expectedValue == null && resource.Type != ProjectResourceType.Work &&
                        (actualValue is decimal zero && zero == 0 || actualValue is ProjectWork || actualValue is ProjectUnits))
                        classification = "xml-omits-inapplicable-resource-field";
                    if (property == "Cost" && totals[resource.Uid].Cost == expected.Cost || property == "ActualCost" && totals[resource.Uid].ActualCost == expected.ActualCost)
                        classification = "xml-resource-total-agrees-with-native-assignment-sum";
                    Compare(actualValue, expectedValue, "resource " + resource.Uid + "/" + property, classification);
                }
                Compare(resource.Calendar?.Uid, expected.Calendar?.Uid, "resource " + resource.Uid + "/Calendar",
                    resource.Uid == 0 && expected.Calendar == null && resource.Calendar?.Uid == 2 ? "xml-omits-reserved-row-calendar" : null);
                Baselines(resource.Baselines, expected.Baselines, "resource " + resource.Uid);
                Custom(resource.CustomFields, expected.CustomFields, "resource " + resource.Uid);
            }
            foreach (var expected in xml.Assignments) {
                var assignment = native.Assignments.SingleOrDefault(a => a.Uid == expected.Uid);
                if (assignment == null) { failures.Add(name + " assignment missing " + expected.Uid); continue; }
                foreach (var property in new[] { "Start", "Finish", "Work", "ActualWork", "RemainingWork", "Cost", "ActualCost", "RemainingCost", "Units" }) {
                    var info = typeof(ProjectAssignment).GetProperty(property)!;
                    Compare(info.GetValue(assignment), info.GetValue(expected), "assignment " + assignment.Uid + "/" + property,
                        property == "Units" && expected.Units == null && assignment.Resource?.Type == ProjectResourceType.Cost ? "xml-omits-cost-resource-units" : null);
                }
                Compare(assignment.Task?.Uid, expected.Task?.Uid, "assignment task " + assignment.Uid);
                Compare(assignment.Resource?.Uid, expected.Resource?.Uid, "assignment resource " + assignment.Uid);
                Baselines(assignment.Baselines, expected.Baselines, "assignment " + assignment.Uid);
            }
            Compare(native.Dependencies.Count, xml.Dependencies.Count, "dependency count");
            foreach (var expected in xml.Dependencies) {
                var actual = native.Dependencies.SingleOrDefault(d => d.Predecessor?.Uid == expected.Predecessor?.Uid && d.Successor.Uid == expected.Successor.Uid);
                string location = "dependency " + expected.Predecessor?.Uid + ":" + expected.Successor.Uid;
                Compare(actual?.Type, expected.Type, location + "/Type"); Compare(actual?.Lag, expected.Lag, location + "/Lag");
                Compare(actual?.LagPercent, expected.LagPercent, location + "/LagPercent");
            }
            Console.WriteLine(name + " tasks=" + native.AllTasks.Count() + " resources=" + native.Resources.Count + " assignments=" + native.Assignments.Count);
        }
        File.WriteAllText(Path.Combine(output, "comparison.json"), JsonSerializer.Serialize(new { comparisons, observations, failures }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("Differences: " + failures.Count);
        foreach (var failure in failures.Take(20)) Console.WriteLine(failure);
        return failures.Count == 0 ? 0 : 1;
    }
}
