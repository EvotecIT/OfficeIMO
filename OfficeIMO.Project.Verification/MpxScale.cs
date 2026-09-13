using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeIMO.Project;

internal static class MpxScale {
    internal static int Run(int count, string mode) {
        if (count < 1 || count > 9000 || (mode != "load" && mode != "cancel")) throw new ArgumentException("Use 1–9000 tasks and load/cancel mode.");
        const int resources = 30;
        var text = new StringBuilder("MPX,Scale fixture,4.0,ANSI\r\n41,40,1\r\n");
        for (int resource = 1; resource <= resources; resource++) text.Append("50,").Append(resource).Append(",Resource ").Append(resource).Append("\r\n");
        text.Append("61,90,1,74\r\n");
        int links = 0;
        for (int task = 1; task <= count; task++) {
            text.Append("70,").Append(task).Append(",Task ").Append(task).Append(",\"");
            for (int predecessor = Math.Max(1, task - 8); predecessor < task; predecessor++) {
                if (predecessor > Math.Max(1, task - 8)) text.Append(',');
                text.Append(predecessor).Append("FS"); links++;
            }
            text.Append("\"\r\n");
            for (int resource = 1; resource <= resources; resource++) text.Append("75,").Append(resource).Append(",1,8h\r\n");
        }
        byte[] input = Encoding.ASCII.GetBytes(text.ToString());
        using var cancellation = new CancellationTokenSource();
        if (mode == "cancel") cancellation.CancelAfter(10);
        long before = GC.GetAllocatedBytesForCurrentThread(); var clock = Stopwatch.StartNew(); bool canceled = false;
        try {
            using var document = ProjectDocument.Load(new MemoryStream(input), cancellationToken: cancellation.Token);
            if (document.Tasks.Count != count || document.Assignments.Count != count * resources || document.Dependencies.Count != links ||
                document.Assignments.Sum(a => a.Work!.Value.Minutes) != count * resources * 480m)
                throw new InvalidDataException("MPX entity, dependency or work totals differ from the generated input.");
        } catch (OperationCanceledException) when (mode == "cancel") { canceled = true; }
        clock.Stop();
        if (mode == "cancel" && (!canceled || clock.ElapsedMilliseconds > 5000)) throw new InvalidDataException("MPX cancellation exceeded its response budget.");
        Console.WriteLine(JsonSerializer.Serialize(new { count, resources, assignments = count * resources, links, mode, inputBytes = input.Length,
            elapsedMs = clock.Elapsed.TotalMilliseconds, allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - before, canceled }));
        return 0;
    }
}
