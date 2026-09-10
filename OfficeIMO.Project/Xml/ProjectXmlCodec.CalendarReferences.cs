namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static void BindCalendars(ProjectDocument document, CancellationToken token) {
        foreach (var calendar in document.Calendars) {
            token.ThrowIfCancellationRequested();
            if (calendar.SourceBaseCalendarUid is int uid && uid > 0 && document.CalendarIndex.TryGetValue(uid, out var parent))
                calendar.BindLoadedBaseCalendar(parent);
        }
        // Each edge is traversed at most once, independent of source record order.
        // The public mutation setter still validates its individual proposed edge.
        var states = new Dictionary<ProjectCalendar, byte>();
        var path = new List<ProjectCalendar>();
        foreach (var calendar in document.Calendars) {
            token.ThrowIfCancellationRequested();
            path.Clear();
            var current = calendar;
            while (current != null) {
                token.ThrowIfCancellationRequested();
                if (states.TryGetValue(current, out byte state)) {
                    if (state == 1) throw new InvalidDataException("Calendar inheritance contains a cycle at UID " + current.Uid + ".");
                    break;
                }
                states.Add(current, 1); path.Add(current); current = current.BaseCalendar;
            }
            foreach (var member in path) { token.ThrowIfCancellationRequested(); states[member] = 2; }
        }
    }
}
