using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static readonly string[] DayOrder = "DayType DayWorking TimePeriod WorkingTimes".Split(' ');
    private static readonly string[] ExceptionOrder = "EnteredByOccurrences TimePeriod Occurrences Name Type Period DaysOfWeek MonthItem MonthPosition Month DayWorking WorkingTimes".Split(' ');
    private static void ReadCalendar(ProjectCalendar calendar, XElement element, CancellationToken token) {
        var ns = element.Name.Namespace;
        calendar.Name = (string?)element.Element(ns + "Name");
        calendar.Guid = element.Element(ns + "GUID") is XElement guid ? ProjectXmlValue.ParseGuid(guid.Value) : (Guid?)null;
        calendar.IsBaseCalendar = (bool?)element.Element(ns + "IsBaseCalendar");
        calendar.SourceBaseCalendarUid = (int?)element.Element(ns + "BaseCalendarUID");
        foreach (var day in Children(element, "WeekDays", "WeekDay")) {
            token.ThrowIfCancellationRequested();
            var item = calendar.WeekDays.Add(); Attach(calendar.Document, item, day);
            int? dayType = (int?)day.Element(ns + "DayType");
            item.Day = dayType >= 1 && dayType <= 7 ? (DayOfWeek?)(dayType - 1) : null;
            item.IsWorking = (bool?)day.Element(ns + "DayWorking");
            var legacyPeriod = day.Element(ns + "TimePeriod");
            item.FromDate = legacyPeriod?.Element(ns + "FromDate") is XElement legacyFrom ? ProjectXmlValue.ParseDate(legacyFrom.Value) : (DateTime?)null;
            item.ToDate = legacyPeriod?.Element(ns + "ToDate") is XElement legacyTo ? ProjectXmlValue.ParseDate(legacyTo.Value) : (DateTime?)null;
            ReadWorkingTimes(item.WorkingTimes, day, calendar.Document, token);
        }
        var legacyByPattern = new Dictionary<string, Queue<ProjectWeekDay>>(StringComparer.Ordinal);
        foreach (var day in calendar.WeekDays.Where(day => day.Day == null)) {
            token.ThrowIfCancellationRequested();
            string key = CalendarPattern(day.FromDate, day.ToDate, day.IsWorking, day.WorkingTimes, token);
            if (!legacyByPattern.TryGetValue(key, out var queue)) { queue = new Queue<ProjectWeekDay>(); legacyByPattern.Add(key, queue); }
            queue.Enqueue(day);
        }
        var mirrored = new HashSet<ProjectWeekDay>();
        foreach (var exception in Children(element, "Exceptions", "Exception")) {
            token.ThrowIfCancellationRequested();
            var item = calendar.Exceptions.Add(); Attach(calendar.Document, item, exception);
            item.Name = (string?)exception.Element(ns + "Name");
            item.IsWorking = (bool?)exception.Element(ns + "DayWorking");
            var period = exception.Element(ns + "TimePeriod");
            item.FromDate = period?.Element(ns + "FromDate") is XElement from ? ProjectXmlValue.ParseDate(from.Value) : (DateTime?)null;
            item.ToDate = period?.Element(ns + "ToDate") is XElement to ? ProjectXmlValue.ParseDate(to.Value) : (DateTime?)null;
            ReadWorkingTimes(item.WorkingTimes, exception, calendar.Document, token);
            string key = CalendarPattern(item.FromDate, item.ToDate, item.IsWorking, item.WorkingTimes, token);
            if (legacyByPattern.TryGetValue(key, out var queue) && queue.Count != 0) {
                var mirror = queue.Dequeue();
                calendar.Document.Source!.LegacyCalendarMirrors.Add(item, calendar.Document.Source.Element(mirror)!);
                for (int i = 0; i < item.WorkingTimes.Count; i++) {
                    token.ThrowIfCancellationRequested();
                    calendar.Document.Source.LegacyWorkingIntervals.Add(item.WorkingTimes[i], calendar.Document.Source.Element(mirror.WorkingTimes[i])!);
                }
                mirrored.Add(mirror);
            }
        }
        calendar.WeekDays.Items.RemoveAll(day => mirrored.Contains(day));
        ReadWorkWeeks(calendar, element, token);
    }
    private static void ReadWorkingTimes(ProjectCollection<ProjectWorkingInterval> intervals, XElement element, ProjectDocument document, CancellationToken token) {
        foreach (var time in Children(element, "WorkingTimes", "WorkingTime")) {
            token.ThrowIfCancellationRequested();
            var item = intervals.Add(); Attach(document, item, time);
            item.From = time.Element(time.Name.Namespace + "FromTime") is XElement from ? ProjectXmlValue.ParseClock(from.Value) : (TimeSpan?)null;
            item.To = time.Element(time.Name.Namespace + "ToTime") is XElement to ? ProjectXmlValue.ParseClock(to.Value) : (TimeSpan?)null;
        }
    }
    private static XElement WriteCalendar(ProjectCalendar calendar, ProjectDocument document, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var node = NewNode(document, calendar, "Calendar");
        void Field(string name, string? value) => ProjectXmlFields.Apply(document, calendar, node, name, value, CalendarOrder);
        Field("UID", ProjectXmlValue.Integer(calendar.Uid)); Field("Name", calendar.Name);
        Field("GUID", ProjectXmlValue.Identifier(calendar.Guid));
        Field("IsBaseCalendar", ProjectXmlValue.Boolean(calendar.IsBaseCalendar));
        Field("BaseCalendarUID", ProjectXmlValue.Integer(calendar.BaseCalendar?.Uid ?? (calendar.IsBaseCalendar == true ? -1 : (int?)null)));
        ReplaceContainer(node, "WeekDays", "WeekDay", calendar.WeekDays.Select(day => {
            token.ThrowIfCancellationRequested();
            var result = NewNode(document, day, "WeekDay");
            ProjectXmlFields.Apply(document, day, result, "DayType", ProjectXmlValue.Integer(day.Day.HasValue ? (int)day.Day.Value + 1 : day.FromDate.HasValue ? 0 : (int?)null), DayOrder);
            ProjectXmlFields.Apply(document, day, result, "DayWorking", ProjectXmlValue.Boolean(day.IsWorking), DayOrder);
            WritePeriod(document, day, result, day.FromDate, day.ToDate, DayOrder);
            ReplaceContainer(result, "WorkingTimes", "WorkingTime", day.WorkingTimes.Select(t => WriteWorkingTime(t, document, token)), DayOrder);
            return result;
        }).Concat(calendar.Exceptions.Where(e => document.Source?.LegacyCalendarMirrors.ContainsKey(e) == true).Select(exception => {
            token.ThrowIfCancellationRequested();
            var result = new XElement(document.Source!.LegacyCalendarMirrors[exception]);
            ProjectXmlFields.Apply(document, exception, result, "DayWorking", ProjectXmlValue.Boolean(exception.IsWorking), DayOrder);
            WritePeriod(document, exception, result, exception.FromDate, exception.ToDate, DayOrder);
            ReplaceContainer(result, "WorkingTimes", "WorkingTime", exception.WorkingTimes.Select(t => WriteWorkingTime(t, document, token, legacy: true)), DayOrder);
            return result;
        })), CalendarOrder);
        ReplaceContainer(node, "Exceptions", "Exception", calendar.Exceptions.Select(exception => {
            token.ThrowIfCancellationRequested();
            var result = NewNode(document, exception, "Exception");
            ProjectXmlFields.Apply(document, exception, result, "Name", exception.Name, ExceptionOrder);
            ProjectXmlFields.Apply(document, exception, result, "DayWorking", ProjectXmlValue.Boolean(exception.IsWorking), ExceptionOrder);
            WritePeriod(document, exception, result, exception.FromDate, exception.ToDate, ExceptionOrder);
            if (document.Source?.Element(exception) == null) {
                ProjectXmlFields.Apply(document, exception, result, "Type", "1", ExceptionOrder);
                ProjectXmlFields.Apply(document, exception, result, "EnteredByOccurrences", "0", ExceptionOrder);
            }
            ReplaceContainer(result, "WorkingTimes", "WorkingTime", exception.WorkingTimes.Select(t => WriteWorkingTime(t, document, token)), ExceptionOrder);
            return result;
        }), CalendarOrder);
        WriteWorkWeeks(calendar, node, token);
        return node;
    }
    private static void WritePeriod(ProjectDocument document, ProjectObject model, XElement result, DateTime? from, DateTime? to, string[] order) {
        var period = result.Element(result.Name.Namespace + "TimePeriod");
        if (period == null && (from.HasValue || to.HasValue)) {
            period = new XElement(result.Name.Namespace + "TimePeriod"); ProjectXmlFields.Insert(result, period, order);
        }
        if (period != null) {
            ProjectXmlFields.Apply(document, model, period, "FromDate", ProjectXmlValue.Date(from), new[] { "FromDate", "ToDate" });
            ProjectXmlFields.Apply(document, model, period, "ToDate", ProjectXmlValue.Date(to), new[] { "FromDate", "ToDate" });
        }
    }
    private static XElement WriteWorkingTime(ProjectWorkingInterval interval, ProjectDocument document, CancellationToken token, bool legacy = false) {
        token.ThrowIfCancellationRequested();
        var result = legacy && document.Source != null && document.Source.LegacyWorkingIntervals.TryGetValue(interval, out var original)
            ? new XElement(original) : NewNode(document, interval, "WorkingTime");
        var order = new[] { "FromTime", "ToTime" };
        ProjectXmlFields.Apply(document, interval, result, "FromTime", ProjectXmlValue.Clock(interval.From), order);
        ProjectXmlFields.Apply(document, interval, result, "ToTime", ProjectXmlValue.Clock(interval.To), order);
        return result;
    }

    private static void ReadDependency(ProjectDependency link, XElement element) {
        var ns = element.Name.Namespace;
        link.Type = element.Element(ns + "Type") is XElement type ? (ProjectDependencyType?)ProjectXmlValue.ParseInt(type.Value) : null;
        link.CrossProject = (bool?)element.Element(ns + "CrossProject");
        link.CrossProjectName = (string?)element.Element(ns + "CrossProjectName");
        int? format = (int?)element.Element(ns + "LagFormat");
        decimal? raw = element.Element(ns + "LinkLag") is XElement lag ? ProjectXmlValue.ParseNumber(lag.Value) : (decimal?)null;
        if (raw.HasValue) {
            if (format == 19 || format == 20 || format == 51 || format == 52) link.LagPercent = raw.Value;
            else link.Lag = ProjectXmlValue.ParseDuration(ProjectXmlValue.Span(ProjectXmlValue.MinutesToSpan(raw.Value / 10)), format, link.Document);
        }
    }
    private static XElement WriteDependency(ProjectDependency link, ProjectDocument document, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var node = NewNode(document, link, "PredecessorLink");
        var order = "PredecessorUID Type CrossProject CrossProjectName LinkLag LagFormat".Split(' ');
        void Field(string name, string? value) => ProjectXmlFields.Apply(document, link, node, name, value, order);
        Field("PredecessorUID", ProjectXmlValue.Integer(link.Predecessor?.Uid ?? link.SourcePredecessorUid));
        Field("Type", ProjectXmlValue.Integer(link.Type.HasValue ? (int?)link.Type.Value : null));
        Field("CrossProject", ProjectXmlValue.Boolean(link.CrossProject)); Field("CrossProjectName", link.CrossProjectName);
        decimal? rawLag = DependencyLag(link, document);
        Field("LinkLag", rawLag.HasValue ? ProjectXmlValue.Number(decimal.Round(rawLag.Value, 0, MidpointRounding.AwayFromZero)) : null);
        Field("LagFormat", link.LagPercent.HasValue ? "19" : link.Lag.HasValue ? ProjectXmlValue.Integer(ProjectXmlValue.DurationFormat(link.Lag.Value)) : null);
        return node;
    }
    internal static decimal? DependencyLag(ProjectDependency link, ProjectDocument document) => link.LagPercent ??
        (link.Lag.HasValue ? checked(link.Lag.Value.Value * ProjectXmlValue.MinutesPerUnit(link.Lag.Value.Unit, link.Lag.Value.IsElapsed, document) * 10) : (decimal?)null);
}
