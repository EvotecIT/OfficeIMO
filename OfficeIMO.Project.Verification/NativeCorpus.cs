using OfficeIMO.Project;
using System.Text.Json;

internal static class NativeCorpus {
    internal static int Run(string input, string output, string pattern = "*.mpp") {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var failures = new List<string>(); var observations = new List<object>(); int comparisons = 0;
        foreach (var path in File.Exists(input) ? new[] { input } : Directory.GetFiles(input, pattern).Where(p => new[] { ".mpp", ".mpt", ".mpx" }.Contains(Path.GetExtension(p).ToLowerInvariant()))) {
            string name = Path.GetFileName(path);
            using var native = ProjectDocument.Load(path);
            using var xml = ProjectDocument.Load(Path.ChangeExtension(path, ".xml"));
            var totals = native.AnalyzeAssignments().Resources.ToDictionary(r => r.ResourceUid);
            using var saved = new MemoryStream(); native.Save(saved);
            if (!saved.ToArray().SequenceEqual(File.ReadAllBytes(path))) failures.Add(name + ": unchanged bytes differ");
            void Compare(object? actual, object? expected, string field, string? representation = null) {
                comparisons++;
                if (Equals(actual, expected)) return;
                if (representation == null && (native.NativeInfo != null || native.MpxSource != null)) {
                    string property = field.Substring(field.LastIndexOf('/') + 1);
                    bool scalar = new[] { "Work", "ActualWork", "RemainingWork", "Cost", "ActualCost", "RemainingCost", "FixedCost", "StandardRate", "OvertimeRate", "CostPerUse",
                        "OvertimeWork", "ActualOvertimeWork", "PercentComplete", "PercentWorkComplete", "FreeSlackMinutes", "TotalSlackMinutes" }.Contains(property);
                    bool flag = new[] { "IsMilestone", "IsCritical", "EffortDriven" }.Contains(property);
                    bool Default(object? value) => value is decimal number && number == 0 || value is int integer && integer == 0 || value is ProjectWork work && work.Minutes == 0 || flag && value is bool boolean && !boolean;
                    if ((scalar || flag) && expected == null && Default(actual)) representation = "xml-omits-explicit-default";
                    else if ((scalar || flag) && actual == null && Default(expected)) representation = "independent-reader-supplies-implicit-default";
                }
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
                    Duration(actual.Duration, baseline.Duration, location + "/Duration", true);
                }
            }
            void Duration(ProjectDuration? actual, ProjectDuration? expected, string field, bool baseline = false) {
                if (!actual.HasValue || !expected.HasValue) {
                    Compare(actual, expected, field, actual == null && expected?.Value == 0 || expected == null && actual?.Value == 0 ? "duration-implicit-zero" : null); return;
                }
                decimal actualMinutes = actual.Value.Value * ProjectTimeUnits.MinutesPerUnit(actual.Value.Unit, actual.Value.IsElapsed, native.Settings);
                decimal expectedMinutes = expected.Value.Value * ProjectTimeUnits.MinutesPerUnit(expected.Value.Unit, expected.Value.IsElapsed, xml.Settings);
                Compare(actualMinutes, expectedMinutes, field + "/Minutes"); Compare(actual.Value.IsElapsed, expected.Value.IsElapsed, field + "/IsElapsed");
                Compare(actual.Value.IsEstimated, expected.Value.IsEstimated, field + "/IsEstimated", baseline && actual.Value.IsEstimated && !expected.Value.IsEstimated ? "xml-baseline-omits-estimate-marker" : null);
                Compare(actual.Value.Unit, expected.Value.Unit, field + "/Unit", actualMinutes == expectedMinutes ? "equivalent-duration-display-units" : null);
            }
            void Custom(IEnumerable<ProjectCustomFieldValue> actuals, IEnumerable<ProjectCustomFieldValue> expecteds, string owner) {
                foreach (var expected in expecteds.Where(v => v.Value != null && v.ValueId == null && v.ValueGuid == null)) {
                    var actual = actuals.SingleOrDefault(v => v.FieldId == expected.FieldId);
                    if (native.MpxSource != null && actual?.Value != null) {
                        var mapping = ProjectMpxFields.CustomMappings(true).Concat(ProjectMpxFields.CustomMappings(false)).FirstOrDefault(f => f.FieldId == expected.FieldId);
                        if (mapping?.Kind == "Duration") {
                            Compare(System.Xml.XmlConvert.ToTimeSpan(actual.Value), System.Xml.XmlConvert.ToTimeSpan(expected.Value!), owner + "/Custom[" + expected.FieldId + "]");
                            continue;
                        }
                        if (mapping?.Kind == "Number" || mapping?.Kind == "Cost") {
                            Compare(decimal.Parse(actual.Value, System.Globalization.CultureInfo.InvariantCulture), decimal.Parse(expected.Value!, System.Globalization.CultureInfo.InvariantCulture), owner + "/Custom[" + expected.FieldId + "]");
                            continue;
                        }
                    }
                    // Project exports enterprise flags in these synthetic local-field fixtures.
                    // Enterprise custom-field records are outside the local scalar profile.
                    Compare(actual?.Value, expected.Value, owner + "/Custom[" + expected.FieldId + "]",
                        actual == null && (expected.FieldId == "188744339" || expected.FieldId == "205521403") ? "unqualified-enterprise-field-preserved-in-native-source" : null);
                }
            }
            foreach (var expected in xml.Calendars) {
                var actual = native.Calendars.SingleOrDefault(c => c.Uid == expected.Uid);
                if (actual == null) { failures.Add(name + " calendar missing " + expected.Uid); continue; }
                var dates = new SortedSet<DateTime>();
                void Include(DateTime from, DateTime to) {
                    if ((to.Date - from.Date).TotalDays > 10000) throw new InvalidDataException("Calendar comparison exceeds its 10000-day fixture budget.");
                    for (var date = from.Date; date <= to.Date; date = date.AddDays(1)) dates.Add(date);
                }
                Include(new DateTime(2000, 1, 3), new DateTime(2000, 1, 16));
                if (xml.Settings.StartDate.HasValue) Include(xml.Settings.StartDate.Value, xml.Settings.StartDate.Value.AddDays(13));
                foreach (var calendar in new[] { actual, expected })
                    for (var current = calendar; current != null; current = current.BaseCalendar) {
                        foreach (var exception in current.Exceptions.Where(e => e.FromDate.HasValue && e.ToDate.HasValue))
                            Include(exception.FromDate!.Value.AddDays(-1), exception.ToDate!.Value.AddDays(1));
                        foreach (var week in current.WorkWeeks.Where(w => w.FromDate.HasValue && w.ToDate.HasValue))
                            Include(week.FromDate!.Value.AddDays(-1), week.ToDate!.Value.AddDays(1));
                    }
                foreach (var date in dates) {
                    string Intervals(ProjectCalendar calendar) => string.Join(";", calendar.GetWorkingIntervals(date)
                        .Select(r => r.Start.ToString("O") + "/" + r.Finish.ToString("O")));
                    Compare(Intervals(actual), Intervals(expected), "calendar " + expected.Uid + "/" + date.ToString("yyyy-MM-dd"));
                }
            }
            foreach (var expected in xml.AllTasks) {
                var task = native.AllTasks.SingleOrDefault(t => t.Uid == expected.Uid);
                if (task == null) { failures.Add(name + " task missing " + expected.Uid); continue; }
                foreach (var property in new[] { "Name", "DisplayId", "Start", "Finish", "EarlyStart", "EarlyFinish", "LateStart", "LateFinish", "ActualStart", "ActualFinish",
                    "Work", "ActualWork", "RemainingWork", "Cost", "ActualCost", "FixedCost", "RemainingCost", "PercentComplete", "PercentWorkComplete", "IsManual", "IsActive",
                    "IsMilestone", "IsSummary", "IsCritical", "ConstraintType", "ConstraintDate", "Type", "Deadline", "Priority", "EffortDriven" }) {
                    var info = typeof(ProjectTask).GetProperty(property)!;
                    string? representation = null;
                    if (native.MpxSource != null && property == "IsCritical" && task.IsCritical == null && expected.IsCritical.HasValue)
                        representation = "independent-reader-supplies-omitted-critical-flag";
                    if (property == "IsSummary" && task.Uid == 0 && task.Children.Count == 0 && task.IsSummary && !expected.IsSummary)
                        representation = "independent-reader-suppresses-empty-project-summary";
                    // Independent exports derive this flag from zero/absent slack. The
                    // installed application also recalculates it on the authored fixture;
                    // retain the differing stored flag as evidence, never silently change it.
                    if (native.NativeInfo != null && property == "IsCritical" && task.IsCritical == false && expected.IsCritical == true &&
                        (task.TotalSlackMinutes ?? 0) == 0 && (task.PercentComplete ?? 0) < 100)
                        representation = "independent-reader-derives-critical-from-zero-slack";
                    if (native.NativeInfo != null && task.Uid == 0 && task.SourceOutlineLevel == 0 && info.GetValue(task) == null &&
                        (property == "ConstraintType" && expected.ConstraintType == ProjectConstraintType.AsSoonAsPossible ||
                         property == "Type" && expected.Type == ProjectTaskType.FixedUnits || property == "Priority" && expected.Priority == 0))
                        representation = "independent-reader-supplies-reserved-summary-default";
                    Compare(info.GetValue(task), info.GetValue(expected), "task " + task.Uid + "/" + property, representation);
                }
                Compare(task.Parent?.Uid, expected.Parent?.Uid, "task " + task.Uid + "/Parent");
                Duration(task.Duration, expected.Duration, "task " + task.Uid + "/Duration");
                Duration(task.ActualDuration, expected.ActualDuration, "task " + task.Uid + "/ActualDuration");
                decimal Minutes(ProjectDuration? value) => value.HasValue ? value.Value.Value * ProjectTimeUnits.MinutesPerUnit(value.Value.Unit, value.Value.IsElapsed, native.Settings) : 0;
                if (native.MpxSource != null && task.RemainingDuration == null && task.Duration.HasValue && expected.RemainingDuration.HasValue &&
                    Minutes(expected.RemainingDuration) == Minutes(task.Duration) - Minutes(task.ActualDuration))
                    Compare(task.RemainingDuration, expected.RemainingDuration, "task " + task.Uid + "/RemainingDuration", "independent-reader-derives-omitted-remaining-duration");
                else if (native.MpxSource != null && task.RemainingDuration == null && Minutes(task.Duration) > 0 && expected.RemainingDuration.HasValue &&
                    task.ActualDuration == null && task.PercentComplete.HasValue &&
                    native.ReadDiagnostics.Any(d => d.Code == "PROJECT_MPX_UNMODELED" && d.Message.StartsWith("Fractional progress", StringComparison.Ordinal)) &&
                    decimal.Round(100m * (1m - Minutes(expected.RemainingDuration) / Minutes(task.Duration)), 0, MidpointRounding.AwayFromZero) == task.PercentComplete.Value)
                    Compare(task.RemainingDuration, expected.RemainingDuration, "task " + task.Uid + "/RemainingDuration", "independent-reader-derives-from-retained-fractional-progress");
                else Duration(task.RemainingDuration, expected.RemainingDuration, "task " + task.Uid + "/RemainingDuration");
                if (native.MpxSource != null) {
                    Compare(task.Notes, expected.Notes, "task " + task.Uid + "/Notes");
                    Compare(task.Contact, expected.Contact, "task " + task.Uid + "/Contact");
                    string Outline(ProjectTask value) {
                        if (value.Uid == 0 && value.SourceOutlineLevel == 0) return "0";
                        var siblings = (value.Parent?.Children ?? native.Tasks).Where(t => t.SourceOutlineLevel != 0).ToArray();
                        string segment = (Array.IndexOf(siblings, value) + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                        return value.Parent == null ? segment : Outline(value.Parent) + "." + segment;
                    }
                    Compare(task.Wbs, expected.Wbs, "task " + task.Uid + "/Wbs", task.Wbs == null && expected.Wbs == Outline(task) ? "independent-reader-derives-omitted-outline-wbs" : null);
                }
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
                if (assignment == null && native.NativeInfo?.Profile.Version == 9 && expected.SourceResourceUid == -65535 && expected.Resource == null) {
                    Compare(null, expected.Uid, "assignment " + expected.Uid, "independent-reader-synthesizes-orphan-assignment-to-reserved-resource");
                    continue;
                }
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
