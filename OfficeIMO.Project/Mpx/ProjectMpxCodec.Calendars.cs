namespace OfficeIMO.Project;

internal static partial class ProjectMpxCodec {
    private sealed partial class Reader {
        private static readonly ProjectWorkingTime[] StandardHours = { ProjectWorkingTime.Hours(8, 12), ProjectWorkingTime.Hours(13, 17) };
        private ProjectCalendar FindCalendar(string name) {
            if (_calendars.TryGetValue(name, out var calendar)) return calendar;
            if (!string.Equals(name, "Standard", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("MPX refers to an undefined calendar: " + name);
            Budget(); calendar = Document.Calendars.AddStandardWorkingWeek(); _calendars.Add(name, calendar); return calendar;
        }
        private void BaseCalendar(string[] r) {
            string name = Get(r, 1);
            if (name.Length == 0 || _calendars.ContainsKey(name)) throw new InvalidDataException("An MPX base calendar needs a unique name.");
            if (_calendars.Count >= 250) throw new InvalidDataException("MPX has more than 250 base calendars.");
            Budget(); _calendar = Document.Calendars.Add(name); _calendars.Add(name, _calendar);
            Week(r, false); Tail(r, 9);
        }
        private void ResourceCalendar(string[] r) {
            if (_resource!.Calendar != null && _resource.Calendar.BaseCalendar != null) throw new InvalidDataException("Duplicate MPX resource calendar.");
            var parent = FindCalendar(Get(r, 1).Length != 0 ? r[1] : "Standard");
            Budget(); _calendar = Document.Calendars.Add(_resource.Name ?? "Resource", parent);
            _resource.Calendar = _calendar; Week(r, true); Tail(r, 9);
        }
        private void Week(string[] r, bool resource) {
            for (int day = 0; day < 7; day++) {
                int flag = Has(r, day + 2) ? ProjectMpxValues.Integer(r[day + 2]) : resource ? 2 : day == 0 || day == 6 ? 0 : 1;
                if (flag == 2 && resource) continue;
                if (flag < 0 || flag > 1) throw new InvalidDataException("Invalid MPX calendar working-day flag.");
                _calendar!.SetWorkingDay((DayOfWeek)day, flag == 0 ? Array.Empty<ProjectWorkingTime>() : StandardHours);
            }
        }
        private ProjectWorkingTime[] Times(string[] r, int start, bool defaultIfEmpty) {
            var times = new List<ProjectWorkingTime>();
            for (int i = start; i < start + 6; i += 2) {
                if (!Has(r, i) && !Has(r, i + 1)) continue;
                if (!Has(r, i) || !Has(r, i + 1)) throw new InvalidDataException("An MPX working interval needs both clock endpoints.");
                times.Add(new ProjectWorkingTime(_values.Time(r[i]), _values.Time(r[i + 1])));
            }
            Tail(r, start + 6);
            return times.Count == 0 && defaultIfEmpty ? StandardHours : times.ToArray();
        }
        private void Hours(string[] r) {
            int day = ProjectMpxValues.Integer(Get(r, 1));
            if (day < 1 || day > 7) throw new InvalidDataException("MPX calendar day must be between 1 and 7.");
            var existing = _calendar!.WeekDays.FirstOrDefault(d => d.Day == (DayOfWeek)(day - 1));
            var times = Times(r, 2, existing?.IsWorking == true);
            _calendar.SetWorkingDay((DayOfWeek)(day - 1), times);
        }
        private void Exception(string[] r) {
            if (_calendar!.Exceptions.Count >= 250) throw new InvalidDataException("MPX calendar exceeds 250 exceptions.");
            DateTime from = _values.Date(Get(r, 1)).Date;
            DateTime to = Has(r, 2) ? _values.Date(r[2]).Date : from;
            if (to < from) throw new InvalidDataException("MPX calendar exception ends before it starts.");
            int flag = Has(r, 3) ? ProjectMpxValues.Integer(r[3]) : 0;
            if (flag == 2 && _phase == 50) { Opaque("Resource calendar exception inheritance remains in the original MPX bytes."); return; }
            if (flag != 0 && flag != 1) throw new InvalidDataException("Invalid MPX exception working flag.");
            var exception = _calendar.Exceptions.Add(); exception.FromDate = from; exception.ToDate = to; exception.IsWorking = flag == 1;
            var times = Times(r, 4, flag == 1);
            if (flag == 0 && times.Length != 0) throw new InvalidDataException("A nonworking MPX exception has working times.");
            foreach (var time in times) { var interval = exception.WorkingTimes.Add(); interval.From = time.From; interval.To = time.To; }
        }
    }
}
