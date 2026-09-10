namespace OfficeIMO.Project;

internal static partial class ProjectModelSnapshot {
    internal static Dictionary<string, object?> Capture(ProjectDocument document, CancellationToken token) {
        var values = new ProjectModelValues();
        void Put(string name, object? value) => values["/Project/" + name] = value;
        Put("Name", document.Name); Put("Title", document.Title); Put("Subject", document.Subject); Put("Author", document.Author);
        Put("Manager", document.Manager); Put("Company", document.Company); Put("Guid", document.Guid);
        Fields(document.Settings, "/Settings", values);
        string Entity(ProjectEntity entity, string kind) {
            token.ThrowIfCancellationRequested(); string path = "/" + kind + "[UID=" + entity.Uid + "]";
            values[path + "/Uid"] = entity.Uid; values[path + "/Guid"] = entity.Guid;
            if (entity is ProjectNamedEntity named) values[path + "/Name"] = named.Name;
            return path;
        }
        void Items<T>(IEnumerable<T> items, string path, Action<T, string, ProjectModelValues> visitor) {
            int index = 0;
            foreach (var item in items) { token.ThrowIfCancellationRequested(); visitor(item, path + "[" + index++ + "]", values); }
            values[path + "/Count"] = index == 0 ? null : (object)index;
        }
        void Rich(ProjectCollection<ProjectBaseline> baselines, ProjectCollection<ProjectCustomFieldValue> fields,
            ProjectCollection<ProjectTimephasedValue> timephased, string path) {
            Items(baselines, path + "/Baseline", (item, key, output) => {
                Fields(item, key, output); Items(item.TimephasedData, key + "/Timephased", Fields);
            });
            Items(fields, path + "/Custom", Fields); Items(timephased, path + "/Timephased", Fields);
        }
        void Days(IEnumerable<ProjectWeekDay> days, string path) => Items(days, path + "/Day", (item, key, output) => {
            Fields(item, key, output); Items(item.WorkingTimes, key + "/Time", Fields);
        });
        int position = 0;
        foreach (var task in document.AllTasks) {
            string path = Entity(task, "Task"); Fields(task, path, values);
            values[path + "/Parent"] = task.Parent?.Uid; values[path + "/Position"] = position++; values[path + "/IsSummary"] = task.IsSummary;
            Rich(task.Baselines, task.CustomFields, task.TimephasedData, path);
        }
        foreach (var resource in document.Resources) { string path = Entity(resource, "Resource"); Fields(resource, path, values); Rich(resource.Baselines, resource.CustomFields, resource.TimephasedData, path); }
        foreach (var assignment in document.Assignments) {
            string path = Entity(assignment, "Assignment"); Fields(assignment, path, values);
            values[path + "/Task"] = assignment.Task?.Uid ?? assignment.SourceTaskUid; values[path + "/Resource"] = assignment.Resource?.Uid ?? assignment.SourceResourceUid;
            Rich(assignment.Baselines, assignment.CustomFields, assignment.TimephasedData, path);
        }
        foreach (var calendar in document.Calendars) {
            string path = Entity(calendar, "Calendar"); Fields(calendar, path, values); values[path + "/BaseCalendar"] = calendar.BaseCalendar?.Uid;
            Days(calendar.WeekDays, path);
            Items(calendar.Exceptions, path + "/Exception", (item, key, output) => { Fields(item, key, output); Items(item.WorkingTimes, key + "/Time", Fields); });
            Items(calendar.WorkWeeks, path + "/Week", (item, key, output) => { Fields(item, key, output); Days(item.WeekDays, key); });
        }
        Items(document.Dependencies, "/Dependency", (item, key, output) => {
            Fields(item, key, output); output[key + "/Predecessor"] = item.Predecessor?.Uid ?? item.SourcePredecessorUid; output[key + "/Successor"] = item.Successor.Uid;
        });
        Items(document.CustomFields, "/Definition", (item, key, output) => { Fields(item, key, output); Items(item.LookupValues, key + "/Lookup", Fields); });
        return values;
    }
}
