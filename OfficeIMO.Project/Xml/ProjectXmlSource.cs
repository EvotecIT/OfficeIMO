using System.Xml.Linq;

namespace OfficeIMO.Project;

/// <summary>Original XML and normalized field snapshots, kept separate from the public semantic model.</summary>
internal sealed class ProjectXmlSource {
    private readonly Dictionary<object, XElement> _elements = new Dictionary<object, XElement>();
    private readonly Dictionary<object, Dictionary<string, string?>> _values = new Dictionary<object, Dictionary<string, string?>>();
    private readonly Dictionary<Type, HashSet<string>> _mappedFields = new Dictionary<Type, HashSet<string>>();
    internal ProjectXmlSource(XDocument xml, byte[]? bytes) { Xml = xml; OriginalBytes = bytes; }
    internal XDocument Xml { get; }
    internal byte[]? OriginalBytes { get; }
    internal int? SaveVersion { get; set; }
    internal string NamespaceName => Xml.Root!.Name.NamespaceName;
    internal bool HasOpaqueStructures { get; set; }
    internal bool Capturing { get; set; }
    internal readonly Dictionary<ProjectCalendarException, XElement> LegacyCalendarMirrors = new Dictionary<ProjectCalendarException, XElement>();
    internal readonly Dictionary<ProjectWorkingInterval, XElement> LegacyWorkingIntervals = new Dictionary<ProjectWorkingInterval, XElement>();
    internal IEnumerable<KeyValuePair<object, XElement>> Elements => _elements;
    internal void Attach(object model, XElement element) => _elements.Add(model, element);
    internal XElement? Element(object model) => _elements.TryGetValue(model, out var element) ? element : null;
    internal XElement CloneOrCreate(object model, string name) => !Capturing && Element(model) is XElement original ? CloneForWrite(original) : new XElement(XName.Get(name, NamespaceName));

    private static XElement CloneForWrite(XElement original) {
        string[] repeated = original.Name.LocalName switch {
            "Task" => new[] { "PredecessorLink", "ExtendedAttribute", "Baseline", "TimephasedData" },
            "Resource" or "Assignment" => new[] { "ExtendedAttribute", "Baseline", "TimephasedData" },
            "Baseline" => new[] { "TimephasedData" },
            "WorkWeek" => new[] { "WeekDay" },
            _ => Array.Empty<string>()
        };
        var clone = new XElement(original.Name, original.Attributes());
        foreach (var node in original.Nodes()) {
            if (node is not XElement child || child.Name.Namespace != original.Name.Namespace) { clone.Add(node); continue; }
            if (repeated.Contains(child.Name.LocalName)) {
                clone.Add(new XElement(child.Name));
                continue;
            }
            string? entry = (original.Name.LocalName, child.Name.LocalName) switch {
                ("Project", "Tasks") => "Task",
                ("Project", "Resources") => "Resource",
                ("Project", "Assignments") => "Assignment",
                ("Project", "Calendars") => "Calendar",
                ("Project", "ExtendedAttributes") => "ExtendedAttribute",
                ("Calendar", "WeekDays") => "WeekDay",
                ("Calendar", "Exceptions") => "Exception",
                ("Calendar", "WorkWeeks") => "WorkWeek",
                ("WorkWeek", "WeekDays") => "WeekDay",
                ("WeekDay", "WorkingTimes") or ("Exception", "WorkingTimes") => "WorkingTime",
                ("ExtendedAttribute", "ValueList") => "Value",
                _ => null
            };
            if (entry == null) { clone.Add(child); continue; }
            var container = new XElement(child.Name, child.Attributes());
            foreach (var member in child.Nodes()) {
                if (member is XElement element && element.Name == child.Name.Namespace + entry) {
                    container.Add(new XElement(element.Name));
                } else container.Add(member);
            }
            clone.Add(container);
        }
        return clone;
    }
    internal void Snapshot(object model, string field, string? value) {
        if (!_mappedFields.TryGetValue(model.GetType(), out var mapped)) { mapped = new HashSet<string>(StringComparer.Ordinal); _mappedFields.Add(model.GetType(), mapped); }
        mapped.Add(field);
        // Most task/resource fields are absent. Their semantic snapshot is null; retain
        // that fact through source membership instead of allocating a dictionary entry
        // for every absent scalar on every entity.
        if (value == null) return;
        if (!_values.TryGetValue(model, out var fields)) { fields = new Dictionary<string, string?>(StringComparer.Ordinal); _values.Add(model, fields); }
        fields[field] = value;
    }
    internal bool Unchanged(object model, string field, string? value) {
        if (!_elements.ContainsKey(model)) return false;
        string? original = null;
        if (_values.TryGetValue(model, out var fields)) fields.TryGetValue(field, out original);
        return string.Equals(original, value, StringComparison.Ordinal);
    }
    internal IEnumerable<string> MappedFields(object model) => _mappedFields.TryGetValue(model.GetType(), out var fields) ? fields : Enumerable.Empty<string>();
}
