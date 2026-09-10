using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Project;

internal static class ScheduleScale {
    internal static int Run(int count, string mode) {
        if (count < 1 || count > 100000 || (mode != "dense" && mode != "cancel")) throw new ArgumentException("Use 1–100000 tasks and dense/cancel mode.");
        using var document = ProjectDocument.Create();
        var start = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = start;
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var tasks = new List<ProjectTask>();
        for (int i = 0; i < count; i++) {
            var task = document.Tasks.Add("Task " + i); task.Duration = ProjectDuration.WorkingMinutes(1);
            for (int preceding = Math.Max(0, i - 4); preceding < i; preceding++) document.Dependencies.Add(tasks[preceding], task);
            tasks.Add(task);
        }
        long revision = document.Revision;
        using var cancellation = new CancellationTokenSource();
        if (mode == "cancel") cancellation.CancelAfter(10);
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        var clock = Stopwatch.StartNew();
        ProjectScheduleResult? result = null; bool canceled = false;
        try { result = document.CalculateSchedule(cancellationToken: cancellation.Token); }
        catch (OperationCanceledException) when (mode == "cancel") { canceled = true; }
        clock.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        if (mode == "cancel" && (!canceled || clock.ElapsedMilliseconds > 5000)) throw new InvalidDataException("Cancellation did not meet its response budget.");
        if (result != null) {
            result.Report.ThrowIfErrors();
            // Independent closed-form finish for a one-minute chain on the standard 8–12/13–17 week.
            int wholeDays = (count - 1) / 480, withinDay = (count - 1) % 480 + 1;
            var expected = start.Date.AddDays(wholeDays / 5 * 7 + wholeDays % 5).AddHours(8)
                .AddMinutes(withinDay + (withinDay > 240 ? 60 : 0));
            if (result.Tasks.Count != count || result.Tasks.Last().Finish != expected || result.Tasks.Any(t => !t.IsCritical || t.TotalSlackMinutes != 0))
                throw new InvalidDataException("Dense schedule differs from the independent chain expectation.");
        }
        if (document.Revision != revision || tasks.Any(t => t.Start.HasValue)) throw new InvalidDataException("Calculation changed stored values.");
        Console.WriteLine(JsonSerializer.Serialize(new { tasks = count, mode, canceled, elapsedMs = clock.Elapsed.TotalMilliseconds,
            allocatedBytes = allocated, peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64, results = result?.Tasks.Count }));
        return 0;
    }
}
