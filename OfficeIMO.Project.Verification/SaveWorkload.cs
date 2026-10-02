using System.Diagnostics;
using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Core;
using OfficeIMO.Project;

/// <summary>Opt-in assessment/save measurement with input construction and reopen proof outside the measured operation.</summary>
internal static class SaveWorkload {
    internal static int Run(int count, ProjectFileFormat format) {
        if (count < 1 || count > 10000) throw new ArgumentOutOfRangeException(nameof(count));
        if (format != ProjectFileFormat.Xml && format != ProjectFileFormat.Mpp14) throw new ArgumentException("Choose Xml or Mpp14.");
        var options = new ProjectSaveOptions { Format = format, Indent = false, LossPolicy = OfficeConversionLossPolicy.Allow };
        byte[] source;
        using (var authored = ProjectDocument.Create()) {
            var start = new DateTime(2026, 10, 5, 8, 0, 0);
            authored.Settings.StartDate = start;
            authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
            var resource = authored.Resources.AddWork("Engineer"); resource.StandardRate = 60m;
            resource.DisplayId = 1; resource.IsNull = false;
            for (int index = 0; index < count; index++) {
                var task = authored.Tasks.Add("Task " + (index + 1));
                task.DisplayId = index + 1; task.IsManual = false; task.IsActive = true; task.IsNull = false;
                task.IsMilestone = false; task.IsCritical = false; task.EffortDriven = false;
                task.Duration = ProjectDuration.WorkingMinutes(60); task.RemainingDuration = task.Duration;
                task.Start = start; task.Finish = start.AddHours(1); task.Work = ProjectWork.Hours(1);
                var assignment = authored.Assignments.Add(task, resource, ProjectUnits.Percent(100));
                assignment.Work = task.Work; assignment.Start = task.Start; assignment.Finish = task.Finish;
            }
            using var initial = new MemoryStream(); authored.Save(initial, options); source = initial.ToArray();
        }
        using var document = ProjectDocument.Load(new MemoryStream(source));
        document.Tasks.GetByUid(count).Name = "Measured edit";
        using var destination = new MemoryStream();
        long before = GC.GetAllocatedBytesForCurrentThread();
        var watch = Stopwatch.StartNew();
        document.AssessSave(options).ThrowIfErrors();
        document.Save(destination, options);
        watch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        using var reopened = ProjectDocument.Load(new MemoryStream(destination.ToArray()));
        if (reopened.AllTasks.Count() != count || reopened.Assignments.Count != count || reopened.Tasks.GetByUid(count).Name != "Measured edit"
            || reopened.AllTasks.Sum(t => t.Work?.Minutes ?? 0m) != count * 60m || reopened.Assignments.Sum(a => a.Work?.Minutes ?? 0m) != count * 60m)
            throw new InvalidDataException("Assessment/save output failed identity, edited-value or work proof.");
        Console.WriteLine(JsonSerializer.Serialize(new { tasks = count, format = format.ToString(), operationMs = watch.Elapsed.TotalMilliseconds,
            allocatedBytes = allocated, sourceBytes = source.Length, outputBytes = destination.Length,
            peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
            runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription }));
        return 0;
    }
}
