using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static readonly ProjectXmlField<ProjectDocument>[] DocumentFields = {
        new ProjectXmlField<ProjectDocument>("Name", (m,d) => m.Name, (m,v,d,e) => m.Name = v),
        new ProjectXmlField<ProjectDocument>("Title", (m,d) => m.Title, (m,v,d,e) => m.Title = v),
        new ProjectXmlField<ProjectDocument>("Subject", (m,d) => m.Subject, (m,v,d,e) => m.Subject = v),
        new ProjectXmlField<ProjectDocument>("Company", (m,d) => m.Company, (m,v,d,e) => m.Company = v),
        new ProjectXmlField<ProjectDocument>("Manager", (m,d) => m.Manager, (m,v,d,e) => m.Manager = v),
        new ProjectXmlField<ProjectDocument>("Author", (m,d) => m.Author, (m,v,d,e) => m.Author = v),
        new ProjectXmlField<ProjectDocument>("GUID", (m,d) => ProjectXmlValue.Identifier(m.Guid), (m,v,d,e) => m.Guid = ProjectXmlValue.ParseGuid(v))
    };
    private static void ReadDocument(ProjectDocument document, XElement root, ProjectLoadOptions options, CancellationToken token) {
        var ns = root.Name.Namespace;
        document.Source!.SaveVersion = (int?)root.Element(ns + "SaveVersion");
        Attach(document, document, root); Attach(document, document.Settings, root);
        ReadFields(document, root, document, DocumentFields);
        // Nullable settings on a loaded document preserve absence; authoring defaults apply only to new documents.
        document.Settings.MinutesPerDay = null; document.Settings.MinutesPerWeek = null;
        document.Settings.DaysPerMonth = null; document.Settings.ScheduleFromStart = null;
        ReadFields(document.Settings, root, document, ProjectXmlFields.Settings);
        document.Settings.SourceCalendarUid = (int?)root.Element(ns + "CalendarUID");
        int entities = 0, timephased = 0;
        foreach (var element in Children(root, "Calendars", "Calendar")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            var calendar = new ProjectCalendar(document, RequiredUid(element));
            if (document.CalendarIndex.ContainsKey(calendar.Uid)) throw new InvalidDataException("Duplicate calendar UID " + calendar.Uid);
            document.Calendars.Items.Add(calendar); document.CalendarIndex.Add(calendar.Uid, calendar); Attach(document, calendar, element);
            ReadCalendar(calendar, element, token);
        }
        BindCalendars(document, token);
        if (document.Settings.SourceCalendarUid is int projectCalendar && document.CalendarIndex.TryGetValue(projectCalendar, out var mainCalendar)) document.Calendar = mainCalendar;
        var levels = new Stack<KeyValuePair<int, ProjectTask>>();
        foreach (var element in Children(root, "Tasks", "Task")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            if (document.TaskIndex.Count >= options.MaxTasks) throw new InvalidDataException("Project exceeds MaxTasks.");
            var task = new ProjectTask(document, RequiredUid(element));
            if (document.TaskIndex.ContainsKey(task.Uid)) throw new InvalidDataException("Duplicate task UID " + task.Uid);
            Attach(document, task, element);
            ReadFields(task, element, document, ProjectXmlFields.Task);
            task.SourceSummary = (bool?)element.Element(ns + "Summary") ?? false;
            task.SourceOutlineLevel = (int?)element.Element(ns + "OutlineLevel");
            int level = task.SourceOutlineLevel ?? (task.Uid == 0 ? 0 : 1);
            if (level < 0 || level > options.MaxOutlineDepth) throw new InvalidDataException("Task outline exceeds MaxOutlineDepth.");
            if ((task.Uid == 0 && level != 0) || (task.Uid != 0 && level == 0))
                throw new InvalidDataException("Only the reserved project summary UID 0 may use outline level zero.");
            while (levels.Count != 0 && levels.Peek().Key >= level) levels.Pop();
            if (levels.Count != 0) {
                if (level != levels.Peek().Key + 1) throw new InvalidDataException("Task outline skips a parent level.");
                task.Parent = levels.Peek().Value;
            } else if (level > 1) throw new InvalidDataException("Task outline has no parent.");
            (task.Parent?.Children ?? document.Tasks).Items.Add(task);
            document.TaskIndex.Add(task.Uid, task);
            // The reserved summary is a project metadata row, not an outline parent.
            if (task.Uid != 0) levels.Push(new KeyValuePair<int, ProjectTask>(level, task));
            task.SourceCalendarUid = (int?)element.Element(ns + "CalendarUID");
            if (task.SourceCalendarUid is int calendarUid && document.CalendarIndex.TryGetValue(calendarUid, out var calendar)) task.Calendar = calendar;
            ReadRich(task.Baselines, task.CustomFields, task.TimephasedData, element, document, options, ref timephased, token);
        }
        foreach (var element in Children(root, "Resources", "Resource")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            var resource = new ProjectResource(document, RequiredUid(element));
            if (document.ResourceIndex.ContainsKey(resource.Uid)) throw new InvalidDataException("Duplicate resource UID " + resource.Uid);
            document.Resources.Items.Add(resource); document.ResourceIndex.Add(resource.Uid, resource); Attach(document, resource, element);
            ReadFields(resource, element, document, ProjectXmlFields.Resource);
            resource.SourceCalendarUid = (int?)element.Element(ns + "CalendarUID");
            if (resource.SourceCalendarUid is int calendarUid && document.CalendarIndex.TryGetValue(calendarUid, out var calendar)) resource.Calendar = calendar;
            ReadRich(resource.Baselines, resource.CustomFields, resource.TimephasedData, element, document, options, ref timephased, token);
        }
        var assignmentIds = new HashSet<int>();
        foreach (var element in Children(root, "Assignments", "Assignment")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            var assignment = new ProjectAssignment(document, RequiredUid(element));
            if (!assignmentIds.Add(assignment.Uid)) throw new InvalidDataException("Duplicate assignment UID " + assignment.Uid);
            Attach(document, assignment, element); ReadFields(assignment, element, document, ProjectXmlFields.Assignment);
            assignment.SourceTaskUid = (int?)element.Element(ns + "TaskUID") ?? throw new InvalidDataException("Assignment TaskUID is missing.");
            assignment.SourceResourceUid = (int?)element.Element(ns + "ResourceUID") ?? -1;
            if (document.TaskIndex.TryGetValue(assignment.SourceTaskUid, out var task)) assignment.Task = task;
            if (document.ResourceIndex.TryGetValue(assignment.SourceResourceUid, out var resource)) assignment.Resource = resource;
            document.Assignments.Items.Add(assignment);
            document.AssignmentIndex.Add(assignment.Uid, assignment);
            if (assignment.Task != null && assignment.Resource != null) document.AssignmentPairs.Add(ProjectDocument.PairKey(assignment.Task.Uid, assignment.Resource.Uid));
            ReadRich(assignment.Baselines, assignment.CustomFields, assignment.TimephasedData, element, document, options, ref timephased, token);
        }
        foreach (var task in document.AllTasks) {
            token.ThrowIfCancellationRequested();
            foreach (var element in document.Source.Element(task)!.Elements(ns + "PredecessorLink")) {
                var link = new ProjectDependency(document) { Successor = task, SourcePredecessorUid = (int?)element.Element(ns + "PredecessorUID") ?? -1 };
                Attach(document, link, element); ReadDependency(link, element);
                if (link.CrossProject != true && document.TaskIndex.TryGetValue(link.SourcePredecessorUid, out var predecessor)) link.Predecessor = predecessor;
                document.Dependencies.Items.Add(link);
                if (link.Predecessor != null) document.DependencyPairs.Add(ProjectDocument.PairKey(link.Predecessor.Uid, link.Successor.Uid));
            }
        }
        foreach (var element in Children(root, "ExtendedAttributes", "ExtendedAttribute")) {
            token.ThrowIfCancellationRequested();
            var definition = document.CustomFields.Add(); Attach(document, definition, element);
            ReadFields(definition, element, document, ProjectXmlFields.CustomFieldDefinition);
            foreach (var value in Children(element, "ValueList", "Value")) {
                var item = definition.LookupValues.Add(); Attach(document, item, value);
                ReadFields(item, value, document, ProjectXmlFields.LookupValue);
            }
        }
    }
    private static void CheckEntities(int count, ProjectLoadOptions options) {
        if (count > options.MaxEntities) throw new InvalidDataException("Project exceeds MaxEntities.");
    }
    private static void ReadRich(ProjectCollection<ProjectBaseline> baselines, ProjectCollection<ProjectCustomFieldValue> custom,
        ProjectCollection<ProjectTimephasedValue> timephased, XElement parent, ProjectDocument document, ProjectLoadOptions options, ref int intervals, CancellationToken token) {
        foreach (var element in parent.Elements(parent.Name.Namespace + "Baseline")) {
            token.ThrowIfCancellationRequested();
            var baseline = baselines.Add(); Attach(document, baseline, element);
            ReadFields(baseline, element, document, ProjectXmlFields.Baseline);
            ReadTimephased(baseline.TimephasedData, element, document, options, ref intervals, token);
        }
        foreach (var element in parent.Elements(parent.Name.Namespace + "ExtendedAttribute")) {
            token.ThrowIfCancellationRequested();
            var value = custom.Add(); Attach(document, value, element);
            ReadFields(value, element, document, ProjectXmlFields.CustomFieldValue);
        }
        ReadTimephased(timephased, parent, document, options, ref intervals, token);
    }
    private static void ReadTimephased(ProjectCollection<ProjectTimephasedValue> values, XElement parent, ProjectDocument document,
        ProjectLoadOptions options, ref int count, CancellationToken token) {
        foreach (var element in parent.Elements(parent.Name.Namespace + "TimephasedData")) {
            token.ThrowIfCancellationRequested();
            if (++count > options.MaxTimephasedValues) throw new InvalidDataException("Project exceeds MaxTimephasedValues.");
            var item = values.Add(); Attach(document, item, element);
            ReadFields(item, element, document, ProjectXmlFields.TimephasedValue);
        }
    }
}
