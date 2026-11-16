using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// Child process that runs one engine. Each UTF-8 stdin line is a JSON <see cref="WorkerRequest"/>; each result is one
/// stdout line prefixed with <see cref="ResultPrefix"/> so stray library output cannot be mistaken for a result.
/// The timed region covers opening the file and extracting the text of every page, nothing else.
/// </summary>
internal static class ExtractionWorker {
    internal const string ResultPrefix = "@@result ";
    internal const string PartialSuffix = ".partial";
    internal static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    internal static int Run(string engine) {
        Func<string, string> extract = TextEngines.Resolve(engine);
        using var output = new StreamWriter(Console.OpenStandardOutput(), Utf8) { AutoFlush = true };
        using var input = new StreamReader(Console.OpenStandardInput(), Utf8);
        Console.SetOut(TextWriter.Null);
        string? line;
        while ((line = input.ReadLine()) is not null) {
            if (line.Length == 0) {
                continue;
            }
            WorkerRequest request = JsonSerializer.Deserialize<WorkerRequest>(line)
                ?? throw new InvalidDataException("Empty worker request.");
            output.WriteLine(ResultPrefix + JsonSerializer.Serialize(Measure(extract, request.Input, request.TextOutput)));
        }
        return 0;
    }

    private static WorkerResult Measure(Func<string, string> extract, string input, string? textOutput) {
        long start = Stopwatch.GetTimestamp();
        string text;
        try {
            text = extract(input);
        } catch (Exception exception) {
            double failedMs = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            string message = exception.Message.Length > 240 ? exception.Message[..240] : exception.Message;
            return new WorkerResult("error", failedMs, 0, exception.GetType().FullName, message, ThrowSite(exception));
        }
        double ms = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        if (textOutput is not null) {
            try {
                PublishText(textOutput, text);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                return new WorkerResult("error", ms, 0, exception.GetType().FullName, "Could not save extracted text: " + exception.Message, null);
            }
        }
        return new WorkerResult("ok", ms, text.Length, null, null, null);
    }

    /// <summary>
    /// Writes beside the final path and renames only after the write completes, so a timeout or I/O failure never
    /// leaves a partial <c>.txt</c> for parity to compare.
    /// </summary>
    private static void PublishText(string path, string text) {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string partial = path + PartialSuffix;
        try {
            File.WriteAllText(partial, text, Utf8);
            File.Move(partial, path, overwrite: true);
        } finally {
            if (File.Exists(partial)) {
                File.Delete(partial);
            }
        }
    }

    /// <summary>The innermost stack frame, which locates the failing code for robustness triage.</summary>
    private static string? ThrowSite(Exception exception) {
        string? trace = exception.StackTrace;
        if (string.IsNullOrEmpty(trace)) {
            return null;
        }
        string first = trace.Split('\n', 2)[0].Trim();
        if (first.StartsWith("at ", StringComparison.Ordinal)) {
            first = first[3..];
        }
        int location = first.IndexOf(" in ", StringComparison.Ordinal);
        return location > 0 ? first[..location] : first;
    }
}

/// <summary>One extraction request. JSON keeps tabs and newlines in file names inside a single request line.</summary>
internal sealed record WorkerRequest(string Input, string? TextOutput);

internal sealed record WorkerResult(
    string Status,
    double Ms,
    int Chars,
    string? ErrorType,
    string? ErrorMessage,
    string? ErrorSite);
