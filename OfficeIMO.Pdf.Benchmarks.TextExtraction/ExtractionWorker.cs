using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// Child process that runs one engine. Each stdin line is <c>input-path\ttext-output-path</c>; each result is one
/// stdout line prefixed with <see cref="ResultPrefix"/> so stray library output cannot be mistaken for a result.
/// The timed region covers opening the file and extracting the text of every page, nothing else.
/// </summary>
internal static class ExtractionWorker {
    internal const string ResultPrefix = "@@result ";
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    internal static int Run(string engine) {
        Func<string, string> extract = TextEngines.Resolve(engine);
        using var output = new StreamWriter(Console.OpenStandardOutput(), Utf8) { AutoFlush = true };
        Console.SetOut(TextWriter.Null);
        string? line;
        while ((line = Console.In.ReadLine()) is not null) {
            if (line.Length == 0) {
                continue;
            }
            string[] parts = line.Split('\t');
            string input = parts[0];
            string? textOutput = parts.Length > 1 && parts[1].Length > 0 ? parts[1] : null;
            output.WriteLine(ResultPrefix + JsonSerializer.Serialize(Measure(extract, input, textOutput)));
        }
        return 0;
    }

    private static WorkerResult Measure(Func<string, string> extract, string input, string? textOutput) {
        long start = Stopwatch.GetTimestamp();
        try {
            string text = extract(input);
            double ms = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            if (textOutput is not null) {
                Directory.CreateDirectory(Path.GetDirectoryName(textOutput)!);
                File.WriteAllText(textOutput, text, Utf8);
            }
            return new WorkerResult("ok", ms, text.Length, null, null, null);
        } catch (Exception exception) {
            double ms = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            string message = exception.Message.Length > 240 ? exception.Message[..240] : exception.Message;
            return new WorkerResult("error", ms, 0, exception.GetType().FullName, message, ThrowSite(exception));
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

internal sealed record WorkerResult(
    string Status,
    double Ms,
    int Chars,
    string? ErrorType,
    string? ErrorMessage,
    string? ErrorSite);
