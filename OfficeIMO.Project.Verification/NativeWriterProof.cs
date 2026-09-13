using OfficeIMO.Project;
using OfficeIMO;

internal static class NativeWriterProof {
    internal static int Create(string output, ProjectFileFormat format = ProjectFileFormat.Automatic, bool minimal = false) {
        if (File.Exists(output)) throw new IOException("Choose a new output file.");
        using var document = ProjectDocument.Create(); document.Name = "Native authored project"; document.Title = "Native authoring proof";
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.ScheduleFromStart = true;
        var calendar = document.Calendars.Add("Standard"); calendar.IsBaseCalendar = true;
        foreach (DayOfWeek day in Enum.GetValues<DayOfWeek>())
            calendar.SetWorkingDay(day, day == DayOfWeek.Saturday || day == DayOfWeek.Sunday ? Array.Empty<ProjectWorkingTime>() : new[] { ProjectWorkingTime.Hours(8, 12), ProjectWorkingTime.Hours(13, 17) });
        document.Calendar = calendar;
        if (minimal) {
            document.Save(output, new ProjectSaveOptions { Format = format });
            return 0;
        }
        var summary = document.Tasks.Add("Delivery");
        var first = summary.Children.Add("Design café / Łódź / 日本語"); first.Duration = ProjectDuration.WorkingDays(1);
        first.Start = document.Settings.StartDate; first.Finish = first.Start.Value.AddHours(9);
        var second = summary.Children.Add("Build"); second.Duration = ProjectDuration.WorkingDays(2);
        second.Start = new DateTime(2026, 10, 6, 8, 0, 0); second.Finish = new DateTime(2026, 10, 7, 17, 0, 0);
        var resource = document.Resources.AddWork("Engineer"); resource.Type = ProjectResourceType.Work; resource.MaxUnits = ProjectUnits.Percent(100);
        resource.StandardRate = 100; resource.Calendar = document.Calendars.Add("Engineer", calendar);
        var assignment = document.Assignments.Add(first, resource); assignment.Units = ProjectUnits.Percent(100); assignment.Work = ProjectWork.Hours(8);
        assignment.Start = first.Start; assignment.Finish = first.Finish;
        document.Dependencies.Add(first, second);
        var baseline = first.Baselines.Add(); baseline.Number = 0; baseline.Start = first.Start; baseline.Finish = first.Finish; baseline.Duration = first.Duration; baseline.Work = ProjectWork.Hours(8); baseline.Cost = 800;
        {
            var definition = document.CustomFields.Add(); definition.FieldId = "188743731"; definition.FieldName = "Text1"; definition.Alias = "Work area";
            var value = first.CustomFields.Add(); value.FieldId = definition.FieldId; value.Value = "Independent authoring";
        }
        {
            var exception = calendar.Exceptions.Add(); exception.Name = "Maintenance"; exception.FromDate = new DateTime(2026, 10, 12); exception.ToDate = exception.FromDate; exception.IsWorking = false;
            var week = calendar.WorkWeeks.Add(); week.Name = "Four day week"; week.FromDate = new DateTime(2026, 10, 19); week.ToDate = new DateTime(2026, 10, 23); week.SetWorkingDay(DayOfWeek.Friday);
        }
        var options = new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };
        foreach (var diagnostic in document.AssessSave(output, options).Diagnostics) Console.WriteLine(diagnostic.Code + " " + diagnostic.Location + " " + diagnostic.Message);
        document.Save(output, options);
        using var read = ProjectDocument.Load(output);
        Console.WriteLine("Tasks=" + read.AllTasks.Count() + " Resources=" + read.Resources.Count + " Assignments=" + read.Assignments.Count + " Dependencies=" + read.Dependencies.Count);
        return 0;
    }
}
