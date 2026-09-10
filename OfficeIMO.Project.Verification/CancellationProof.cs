using System.Diagnostics;
using OfficeIMO.Project;

internal static class CancellationProof {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        double loadMilliseconds = CheckCancelled(token => {
            using var ignored = ProjectDocument.Load(input, cancellationToken: token);
        });
        using var project = ProjectDocument.Load(input);
        project.Name = "Cancellation proof";
        string target = Path.Combine(output, "retained.xml");
        byte[] sentinel = System.Text.Encoding.UTF8.GetBytes("Existing destination must remain unchanged");
        File.WriteAllBytes(target, sentinel);
        double saveMilliseconds = CheckCancelled(token => project.Save(target, cancellationToken: token));
        if (!File.ReadAllBytes(target).SequenceEqual(sentinel) || !project.IsModified)
            throw new InvalidDataException("Cancelled save changed its destination or accepted the model revision.");
        string report = System.Text.Json.JsonSerializer.Serialize(new {
            cancellationRequestedAfterMs = 25, responseBudgetMs = 5000, loadMilliseconds, saveMilliseconds,
            destinationUnchanged = true, modelRemainsModified = true
        });
        File.WriteAllText(Path.Combine(output, "cancellation.json"), report);
        Console.WriteLine(report);
        return 0;
    }

    private static double CheckCancelled(Action<CancellationToken> operation) {
        using var source = new CancellationTokenSource();
        var elapsed = Stopwatch.StartNew();
        source.CancelAfter(25);
        try { operation(source.Token); }
        catch (OperationCanceledException) when (source.IsCancellationRequested) {
            if (elapsed.ElapsedMilliseconds > 5000) throw new InvalidDataException("Cancellation exceeded the five-second response budget.");
            return elapsed.Elapsed.TotalMilliseconds;
        }
        throw new InvalidDataException("Operation completed before cancellation. Use the 100,000-task scale input for this check.");
    }
}
