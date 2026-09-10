using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static void ReadWorkWeeks(ProjectCalendar calendar, XElement element, CancellationToken token) {
        var ns = element.Name.Namespace;
        foreach (var source in Children(element, "WorkWeeks", "WorkWeek")) {
            token.ThrowIfCancellationRequested();
            var week = calendar.WorkWeeks.Add(); Attach(calendar.Document, week, source);
            week.Name = (string?)source.Element(ns + "Name");
            var period = source.Element(ns + "TimePeriod");
            week.FromDate = period?.Element(ns + "FromDate") is XElement from ? ProjectXmlValue.ParseDate(from.Value) : (DateTime?)null;
            week.ToDate = period?.Element(ns + "ToDate") is XElement to ? ProjectXmlValue.ParseDate(to.Value) : (DateTime?)null;
            foreach (var sourceDay in source.Elements(ns + "WeekDay").Concat(Children(source, "WeekDays", "WeekDay"))) {
                token.ThrowIfCancellationRequested();
                var day = week.WeekDays.Add(); Attach(calendar.Document, day, sourceDay);
                int? type = (int?)sourceDay.Element(ns + "DayType");
                day.Day = type >= 1 && type <= 7 ? (DayOfWeek?)(type - 1) : null;
                day.IsWorking = (bool?)sourceDay.Element(ns + "DayWorking");
                ReadWorkingTimes(day.WorkingTimes, sourceDay, calendar.Document, token);
            }
        }
    }
    private static void WriteWorkWeeks(ProjectCalendar calendar, XElement node, CancellationToken token) {
        var document = calendar.Document;
        var order = new[] { "TimePeriod", "Name", "WeekDays", "WeekDay" };
        ReplaceContainer(node, "WorkWeeks", "WorkWeek", calendar.WorkWeeks.Select(week => {
            token.ThrowIfCancellationRequested();
            var result = NewNode(document, week, "WorkWeek");
            WritePeriod(document, week, result, week.FromDate, week.ToDate, order);
            ProjectXmlFields.Apply(document, week, result, "Name", week.Name, order);
            var days = week.WeekDays.Select(day => {
                token.ThrowIfCancellationRequested();
                var output = NewNode(document, day, "WeekDay");
                ProjectXmlFields.Apply(document, day, output, "DayType", ProjectXmlValue.Integer(day.Day.HasValue ? (int)day.Day.Value + 1 : (int?)null), DayOrder);
                ProjectXmlFields.Apply(document, day, output, "DayWorking", ProjectXmlValue.Boolean(day.IsWorking), DayOrder);
                ReplaceContainer(output, "WorkingTimes", "WorkingTime", day.WorkingTimes.Select(time => WriteWorkingTime(time, document, token)), DayOrder);
                return output;
            }).ToArray();
            bool direct = document.Source?.Element(week)?.Elements(result.Name.Namespace + "WeekDay").Any() == true;
            result.Elements(result.Name.Namespace + "WeekDay").Remove();
            if (direct) foreach (var day in days) ProjectXmlFields.Insert(result, day, order);
            else ReplaceContainer(result, "WeekDays", "WeekDay", days, order);
            return result;
        }), CalendarOrder);
    }
}
