using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    internal static bool RequiresRemainingDurationDefault(ProjectDocument document, ProjectTask task) =>
        document.Source?.Element(task) == null && task.Duration.HasValue && !task.RemainingDuration.HasValue &&
        !task.ActualDuration.HasValue && (task.PercentComplete ?? 0) == 0;

    private static XDocument WriteDocument(ProjectDocument document, CancellationToken token) {
        var xml = new XDocument(new XDeclaration("1.0", "utf-8", document.Source?.Xml.Declaration?.Standalone));
        if (document.Source != null && !document.Source.Capturing) {
            foreach (var node in document.Source.Xml.Nodes()) xml.Add(node == document.Source.Xml.Root ? NewNode(document, document, "Project") : node);
        } else xml.Add(new XElement(XName.Get("Project", document.XmlNamespace)));
        var root = xml.Root!;
        if (document.Source == null) root.Add(new XElement(root.Name.Namespace + "SaveVersion", 14));
        ProjectXmlFields.Write(document, root, document, DocumentFields, RootOrder);
        ProjectXmlFields.Write(document.Settings, root, document, ProjectXmlFields.Settings, RootOrder);
        ProjectXmlFields.Apply(document, document.Settings, root, "CalendarUID", ProjectXmlValue.Integer(document.Calendar?.Uid ?? document.Settings.SourceCalendarUid), RootOrder);
        ReplaceContainer(root, "ExtendedAttributes", "ExtendedAttribute", document.CustomFields.Select(field => WriteDefinition(field, document, token)), RootOrder);
        WriteOutlineCodes(document, root, token);
        ReplaceContainer(root, "Calendars", "Calendar", document.Calendars.Select(calendar => WriteCalendar(calendar, document, token)), RootOrder);
        var links = document.Dependencies.GroupBy(d => d.Successor).ToDictionary(g => g.Key, g => g.ToArray());
        int row = 0;
        var levels = new Dictionary<ProjectTask, int>();
        ReplaceContainer(root, "Tasks", "Task", document.AllTasks.Select(task => {
            token.ThrowIfCancellationRequested();
            var node = NewNode(document, task, "Task");
            ProjectXmlFields.Write(task, node, document, ProjectXmlFields.Task, TaskOrder);
            int level = task.Uid == 0 ? 0 : task.Parent != null ? levels[task.Parent] + 1 : 1;
            levels.Add(task, level);
            void Field(string name, string? value) => ProjectXmlFields.Apply(document, task, node, name, value, TaskOrder);
            Field("UID", ProjectXmlValue.Integer(task.Uid));
            Field("ID", ProjectXmlValue.Integer(document.StructureChanged || document.Source == null ? (task.Uid == 0 ? 0 : ++row) : task.DisplayId));
            Field("OutlineLevel", ProjectXmlValue.Integer(document.StructureChanged || document.Source == null ? level : task.SourceOutlineLevel));
            Field("Summary", ProjectXmlValue.Boolean(task.IsSummary));
            Field("CalendarUID", ProjectXmlValue.Integer(task.Calendar?.Uid ?? task.SourceCalendarUid));
            if (!ProjectXmlValue.TryTaskDurationFormat(task, out int? durationFormat))
                throw new InvalidDataException("XML task durations require one shared unit and flags.");
            Field("DurationFormat", ProjectXmlValue.Integer(durationFormat));
            // Project imports a newly authored, unstarted task as zero duration when
            // RemainingDuration is absent. This initial value is a serialization default,
            // not calendar arithmetic or recalculation of an imported schedule.
            if (RequiresRemainingDurationDefault(document, task))
                Field("RemainingDuration", ProjectXmlValue.Duration(task.Duration, document));
            ReplaceChildren(node, "PredecessorLink", links.TryGetValue(task, out var predecessors) ? predecessors.Select(d => WriteDependency(d, document, token)) : Enumerable.Empty<XElement>(), TaskOrder);
            WriteRich(node, task.Baselines, task.CustomFields, task.TimephasedData, document, TaskOrder, token);
            WriteOutlineSelections(task.OutlineCodes, node, TaskOrder, token);
            return node;
        }), RootOrder);
        ReplaceContainer(root, "Resources", "Resource", document.Resources.Select(resource => {
            token.ThrowIfCancellationRequested();
            var node = NewNode(document, resource, "Resource");
            ProjectXmlFields.Write(resource, node, document, ProjectXmlFields.Resource, ResourceOrder);
            ProjectXmlFields.Apply(document, resource, node, "UID", ProjectXmlValue.Integer(resource.Uid), ResourceOrder);
            ProjectXmlFields.Apply(document, resource, node, "CalendarUID", ProjectXmlValue.Integer(resource.Calendar?.Uid ?? resource.SourceCalendarUid), ResourceOrder);
            WriteRich(node, resource.Baselines, resource.CustomFields, resource.TimephasedData, document, ResourceOrder, token);
            WriteResourceCapacity(resource, node, token);
            WriteOutlineSelections(resource.OutlineCodes, node, ResourceOrder, token);
            return node;
        }), RootOrder);
        ReplaceContainer(root, "Assignments", "Assignment", document.Assignments.Select(assignment => {
            token.ThrowIfCancellationRequested();
            var node = NewNode(document, assignment, "Assignment");
            ProjectXmlFields.Write(assignment, node, document, ProjectXmlFields.Assignment, AssignmentOrder);
            ProjectXmlFields.Apply(document, assignment, node, "UID", ProjectXmlValue.Integer(assignment.Uid), AssignmentOrder);
            ProjectXmlFields.Apply(document, assignment, node, "TaskUID", ProjectXmlValue.Integer(assignment.Task?.Uid ?? assignment.SourceTaskUid), AssignmentOrder);
            ProjectXmlFields.Apply(document, assignment, node, "ResourceUID", ProjectXmlValue.Integer(assignment.Resource?.Uid ?? assignment.SourceResourceUid), AssignmentOrder);
            WriteRich(node, assignment.Baselines, assignment.CustomFields, assignment.TimephasedData, document, AssignmentOrder, token);
            return node;
        }), RootOrder);
        return xml;
    }

    private static void WriteRich(XElement node, ProjectCollection<ProjectBaseline> baselines, ProjectCollection<ProjectCustomFieldValue> fields,
        ProjectCollection<ProjectTimephasedValue> timephased, ProjectDocument document, string[] order, CancellationToken token) {
        ReplaceChildren(node, "ExtendedAttribute", fields.Select(field => {
            token.ThrowIfCancellationRequested();
            var result = NewNode(document, field, "ExtendedAttribute");
            ProjectXmlFields.Write(field, result, document, ProjectXmlFields.CustomFieldValue, new[] { "FieldID", "Value", "ValueID", "ValueGUID", "DurationFormat" });
            return result;
        }), order);
        ReplaceChildren(node, "Baseline", baselines.Select(baseline => {
            token.ThrowIfCancellationRequested();
            var result = NewNode(document, baseline, "Baseline");
            ProjectXmlFields.Write(baseline, result, document, ProjectXmlFields.Baseline, BaselineOrder);
            ProjectXmlFields.Apply(document, baseline, result, "DurationFormat", baseline.Duration.HasValue ? ProjectXmlValue.Integer(ProjectXmlValue.DurationFormat(baseline.Duration.Value)) : null, BaselineOrder);
            ReplaceChildren(result, "TimephasedData", baseline.TimephasedData.Select(t => WriteTimephased(t, document, token)), BaselineOrder);
            return result;
        }), order);
        ReplaceChildren(node, "TimephasedData", timephased.Select(t => WriteTimephased(t, document, token)), order);
    }
    private static XElement WriteTimephased(ProjectTimephasedValue value, ProjectDocument document, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var result = NewNode(document, value, "TimephasedData");
        ProjectXmlFields.Write(value, result, document, ProjectXmlFields.TimephasedValue, new[] { "Type", "UID", "Start", "Finish", "Unit", "Value" });
        return result;
    }
    private static XElement WriteDefinition(ProjectCustomFieldDefinition definition, ProjectDocument document, CancellationToken token) {
        var result = NewNode(document, definition, "ExtendedAttribute");
        var order = "FieldID FieldName CFType Guid ElemType MaxMultiValues UserDef Alias SecondaryPID AutoRollDown DefaultGuid Ltuid SecondaryGuid IsBaseline FieldType RollupType CalculationType Formula RestrictValues ValuelistSortOrder AppendNewValues Default ValueList".Split(' ');
        ProjectXmlFields.Write(definition, result, document, ProjectXmlFields.CustomFieldDefinition, order);
        ReplaceContainer(result, "ValueList", "Value", definition.LookupValues.Select(value => {
            token.ThrowIfCancellationRequested();
            var node = NewNode(document, value, "Value");
            ProjectXmlFields.Write(value, node, document, ProjectXmlFields.LookupValue, new[] { "ID", "Value", "Description", "GUID" });
            return node;
        }), order);
        return result;
    }

    private static void CaptureSnapshots(ProjectDocument document, CancellationToken token) {
        document.Source!.Capturing = true;
        try { _ = WriteDocument(document, token); }
        finally { document.Source.Capturing = false; }
    }

    private static void InspectPreservedContent(ProjectDocument document, ProjectLoadOptions options, CancellationToken token) {
        var source = document.Source!;
        int count = 0;
        var containers = new HashSet<string>("ExtendedAttributes Calendars Tasks Resources Assignments WeekDays WorkingTimes Exceptions WorkWeeks TimePeriod ValueList AvailabilityPeriods Rates".Split(' '), StringComparer.Ordinal);
        var known = new HashSet<XElement>();
        foreach (var pair in source.Elements) {
            token.ThrowIfCancellationRequested();
            known.Add(pair.Value);
            foreach (string field in source.MappedFields(pair.Key)) {
                token.ThrowIfCancellationRequested();
                foreach (var scalar in pair.Value.Elements(pair.Value.Name.Namespace + field)) known.Add(scalar);
                var period = pair.Value.Element(pair.Value.Name.Namespace + "TimePeriod");
                if (period?.Element(period.Name.Namespace + field) is XElement date) known.Add(date);
            }
        }
        var mirrorFields = new HashSet<string>("DayType DayWorking FromDate ToDate FromTime ToTime WorkingTime".Split(' '), StringComparer.Ordinal);
        foreach (var mirror in source.LegacyCalendarMirrors.Values) {
            token.ThrowIfCancellationRequested();
            known.Add(mirror);
            foreach (var node in mirror.Descendants()) {
                token.ThrowIfCancellationRequested();
                if (node.Name.Namespace == mirror.Name.Namespace && mirrorFields.Contains(node.Name.LocalName)) known.Add(node);
            }
        }
        if (source.Xml.Root!.Element(source.Xml.Root.Name.Namespace + "SaveVersion") is XElement version) known.Add(version);
        var pending = new Stack<XElement>(); pending.Push(source.Xml.Root);
        while (pending.Count != 0) {
            token.ThrowIfCancellationRequested();
            var element = pending.Pop();
            bool understood = known.Contains(element) || (element.Name.NamespaceName == source.NamespaceName && containers.Contains(element.Name.LocalName));
            if (!understood || element.Attributes().Any(a => !a.IsNamespaceDeclaration)) {
                source.HasOpaqueStructures = true;
                if (++count <= options.MaxDiagnostics) document.AddReadDiagnostic(new ProjectDiagnostic("PROJECT_XML_PRESERVED", ProjectDiagnosticSeverity.Information,
                    "XML content or attributes are retained but not semantically interpreted; structural edits may invalidate their references.", ProjectXmlValue.Location(element.Parent ?? element) + "/" + element.Name.LocalName));
            }
            if (understood) foreach (var child in element.Elements()) pending.Push(child);
        }
        if (count > options.MaxDiagnostics) document.AddReadDiagnostic(new ProjectDiagnostic("PROJECT_DIAGNOSTICS_TRUNCATED", ProjectDiagnosticSeverity.Warning,
            "Additional preserved-content diagnostics exceeded MaxDiagnostics.", "/Project"));
    }
}
