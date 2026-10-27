using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

internal static class CupsPdfPrinter {
    internal static async Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken token) {
        if (!(await GetPrintersAsync(token).ConfigureAwait(false)).Any(printer => printer.Name == printerName))
            throw new ArgumentException("The selected printer queue is not available.", nameof(printerName));
        var result = await RunAsync("lpoptions", ["-p", printerName, "-l"], token).ConfigureAwait(false);
        if (result.ExitCode != 0) throw new IOException(result.Error.Trim());
        return ParsePaperSources(result.Output);
    }

    internal static IReadOnlyList<PdfPaperSourceInfo> ParsePaperSources(string output) {
        foreach (string line in output.Split('\n')) {
            int separator = line.IndexOf(':');
            if (separator <= 0) continue;
            string key = line.Substring(0, separator).Split('/')[0];
            if (key is not ("InputSlot" or "media-source")) continue;
            var sources = new List<PdfPaperSourceInfo>();
            foreach (string item in line.Substring(separator + 1).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                string choice = item.TrimStart('*');
                if (choice.Length is 0 or > 128 || !choice.All(character => OfficeRuntimeCompatibility.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '.')) continue;
                sources.Add(new(key + "=" + choice, choice));
            }
            return sources.DistinctBy(source => source.Id).ToArray();
        }
        return [];
    }

    internal static async Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken token) {
        var queues = await RunAsync("lpstat", ["-e"], token).ConfigureAwait(false);
        if (queues.ExitCode != 0) throw new IOException(queues.Error.Trim());
        var current = await RunAsync("lpstat", ["-d"], token).ConfigureAwait(false);
        const string prefix = "system default destination: ";
        string? defaultName = current.Output.Trim().StartsWith(prefix, StringComparison.Ordinal)
            ? current.Output.Trim().Substring(prefix.Length) : null;
        return queues.Output.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(name => name.Trim()).Where(name => name.Length != 0)
            .Distinct(StringComparer.Ordinal).Select(name => new PdfPrinterInfo(name, name == defaultName, false))
            .OrderByDescending(printer => printer.IsDefault).ThenBy(printer => printer.Name, StringComparer.Ordinal).ToArray();
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!(options.OutputFilePath == null || string.IsNullOrWhiteSpace(options.OutputFilePath))) throw new NotSupportedException("CUPS file output is controlled by the selected printer queue.");
        if (!(await GetPrintersAsync(token).ConfigureAwait(false)).Any(printer => printer.Name == options.PrinterName))
            throw new ArgumentException("The selected printer queue is not available.", nameof(options));
        if (options.PaperSourceId is not null && !(await GetPaperSourcesAsync(options.PrinterName, token).ConfigureAwait(false)).Any(source => source.Id == options.PaperSourceId))
            throw new ArgumentException("The selected paper source is no longer available on this printer.", nameof(options));
        PdfPageSize paper = document.Sheets[0].Plan.PaperSize;
        double width = Math.Min(paper.Width.ToPoints(), paper.Height.ToPoints()), height = Math.Max(paper.Width.ToPoints(), paper.Height.ToPoints());
        if (document.Sheets.Any(sheet => Math.Abs(Math.Min(sheet.Plan.PaperSize.Width.ToPoints(), sheet.Plan.PaperSize.Height.ToPoints()) - width) > 1 ||
            Math.Abs(Math.Max(sheet.Plan.PaperSize.Width.ToPoints(), sheet.Plan.PaperSize.Height.ToPoints()) - height) > 1))
            throw new NotSupportedException("One CUPS job must use a single paper size.");
        var pages = new List<PdfDocument>();
        foreach (PdfRenderedPrintSheet sheet in document.Sheets) {
            token.ThrowIfCancellationRequested();
            pages.Add(PdfDocument.CreateFromImages([new PdfImageDocumentSource(sheet.Png)],
                new PdfImageImportOptions { FixedPageSize = sheet.Plan.PaperSize, Margin = 0, Fit = OfficeImageFit.Stretch }, token));
        }
        byte[] bytes = SerializeSpool(pages, token);
        return await SubmitSpoolAsync(document, options, bytes, token, RunAsync).ConfigureAwait(false);
    }

    internal static byte[] SerializeSpool(IReadOnlyList<PdfDocument> pages, CancellationToken token) {
        byte[] bytes = (pages.Count == 1 ? pages[0] : PdfDocument.Merge(pages, token)).ToBytes(new PdfSaveOptions { AllowLoss = true }, token);
        if (bytes.LongLength > 256L * 1024 * 1024) throw new InvalidOperationException("The CUPS spool document exceeds its byte limit.");
        return bytes;
    }

    internal delegate Task<(int ExitCode, string Output, string Error)> CupsCommandRunner(
        string command, IReadOnlyList<string> arguments, CancellationToken token, Action? started);

    internal static async Task<PdfPrintSubmission> SubmitSpoolAsync(PdfPreparedPrintDocument document,
        PdfPrintDeliveryOptions options, byte[] bytes, CancellationToken token, CupsCommandRunner run) {
        token.ThrowIfCancellationRequested();
        PdfPageSize paper = document.Sheets[0].Plan.PaperSize;
        double width = Math.Min(paper.Width.ToPoints(), paper.Height.ToPoints()), height = Math.Max(paper.Width.ToPoints(), paper.Height.ToPoints());
        string directory = OfficeIMO.Core.Internal.OfficeTemporaryDirectory.Create("officeimo-print-");
        bool submissionUncertain = false;
        string? jobId = null;
        Exception? failure = null;
        PdfPrintSubmission? submission = null;
        try {
            string path = Path.Combine(directory, "sheets.pdf");
            using (FileStream spool = OfficeIMO.Core.Internal.OfficeTemporaryFile.CreateAtPath(path, 81920, FileOptions.Asynchronous)) {
                await spool.WriteAsync(bytes, token).ConfigureAwait(false);
            }
            var arguments = new List<string> {
                "-d", options.PrinterName, "-n", options.Copies.ToString(CultureInfo.InvariantCulture), "-t", options.DocumentName,
                "-o", "media=" + Media(width, height), "-o", "print-scaling=none", "-o", "number-up=1", "-o", "Collate=True"
            };
            if (document.Sheets.All(sheet => sheet.Plan.PaperSize.Width.ToPoints() > sheet.Plan.PaperSize.Height.ToPoints())) {
                arguments.Add("-o"); arguments.Add("orientation-requested=4");
            } else if (document.Sheets.All(sheet => sheet.Plan.PaperSize.Width.ToPoints() <= sheet.Plan.PaperSize.Height.ToPoints())) {
                arguments.Add("-o"); arguments.Add("orientation-requested=3");
            }
            if (options.Duplex != PdfPrintDuplex.PrinterDefault) {
                arguments.Add("-o");
                arguments.Add("sides=" + (options.Duplex == PdfPrintDuplex.SingleSided ? "one-sided" : options.Duplex == PdfPrintDuplex.LongEdge ? "two-sided-long-edge" : "two-sided-short-edge"));
            }
            if (options.PaperSourceId is not null) {
                arguments.Add("-o"); arguments.Add(options.PaperSourceId);
            }
            arguments.Add("--"); arguments.Add(path);
            var result = await run("lp", arguments, token, () => submissionUncertain = true).ConfigureAwait(false);
            if (result.ExitCode != 0) {
                // A missing destination is diagnosed before lp creates a job. Other exit-1
                // failures can follow an accepted upload and an unsuccessful cancellation,
                // so the exit code alone must never make retrying appear safe.
                if (result.ExitCode == 1 && (result.Output == null || string.IsNullOrWhiteSpace(result.Output)) &&
                    result.Error.Trim() is "lp: Error - The printer or class does not exist." or
                        "/usr/bin/lp: Error - The printer or class does not exist.")
                    submissionUncertain = false;
                throw new IOException((result.Error == null || string.IsNullOrWhiteSpace(result.Error))
                    ? $"The print command failed with exit code {result.ExitCode}." : result.Error.Trim());
            }
            Match receipt = Regex.Match(result.Output, @"request id is (\S+-\d+)", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            if (!receipt.Success) throw new IOException("The print command returned no recognizable job acknowledgement.");
            jobId = receipt.Groups[1].Value;
            submission = new(options.PrinterName, jobId, document.Sheets.Count, options.Copies, null);
        } catch (Exception error) {
            failure = error;
            if (submissionUncertain) throw new PdfPrintDeliveryException(jobId, error);
            throw;
        } finally {
            submission = CleanupStaging(directory, submission, failure, submissionUncertain, jobId, path => Directory.Delete(path, recursive: true));
        }
        return submission!;
    }

    internal static PdfPrintSubmission? CleanupStaging(string directory, PdfPrintSubmission? receipt, Exception? failure,
        bool submissionUncertain, string? jobId, Action<string> delete) {
        try { delete(directory); }
        catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException) {
            string message = $"Could not remove print staging at '{directory}': {cleanup.Message}";
            if (receipt is not null) return receipt with { CleanupWarning = message };
            Exception detail = new IOException(message, cleanup);
            if (failure is not null) detail = new AggregateException(failure, detail);
            if (submissionUncertain) throw new PdfPrintDeliveryException(jobId, detail);
            throw detail;
        }
        return receipt;
    }

    private static string Media(double width, double height) {
        if (Math.Abs(width - PdfPageSizes.A4.Width.ToPoints()) < 1 && Math.Abs(height - PdfPageSizes.A4.Height.ToPoints()) < 1) return "A4";
        if (Math.Abs(width - PdfPageSizes.A3.Width.ToPoints()) < 1 && Math.Abs(height - PdfPageSizes.A3.Height.ToPoints()) < 1) return "A3";
        if (Math.Abs(width - PdfPageSizes.Letter.Width.ToPoints()) < 1 && Math.Abs(height - PdfPageSizes.Letter.Height.ToPoints()) < 1) return "Letter";
        if (Math.Abs(width - PdfPageSizes.Legal.Width.ToPoints()) < 1 && Math.Abs(height - PdfPageSizes.Legal.Height.ToPoints()) < 1) return "Legal";
        return string.Format(CultureInfo.InvariantCulture, "Custom.{0:0.###}x{1:0.###}pt", width, height);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunAsync(string command, IReadOnlyList<string> arguments, CancellationToken token, Action? started = null) {
        string executable = "/usr/bin/" + command;
        if (!File.Exists(executable)) throw new PlatformNotSupportedException("CUPS command-line printing tools are not installed.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        start.Environment["LC_ALL"] = "C";
#if NET8_0_OR_GREATER
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
#else
        start.Arguments = string.Join(" ", arguments.Select(QuoteProcessArgument));
#endif
        using var process = new Process { StartInfo = start };
        timeout.Token.ThrowIfCancellationRequested();
        if (!process.Start()) throw new IOException("The print command could not start.");
        started?.Invoke();
        Task<string> output = ReadBoundedAsync(process.StandardOutput, timeout.Token);
        Task<string> error = ReadBoundedAsync(process.StandardError, timeout.Token);
        try {
            await WaitForExitAsync(process, timeout.Token).ConfigureAwait(false);
            return (process.ExitCode, await output.ConfigureAwait(false), await error.ConfigureAwait(false));
        } catch {
            try {
                if (!process.HasExited) {
#if NET8_0_OR_GREATER
                    process.Kill(entireProcessTree: true);
#else
                    process.Kill();
#endif
                }
            } catch (InvalidOperationException) { }
            await WaitForExitAsync(process, CancellationToken.None).ConfigureAwait(false);
            try { await Task.WhenAll(output, error).ConfigureAwait(false); }
            catch (Exception readError) when (readError is IOException or OperationCanceledException) { }
            throw;
        }
    }

    private static async Task<string> ReadBoundedAsync(StreamReader reader, CancellationToken token) {
        var text = new StringBuilder();
        char[] buffer = new char[4096];
        bool exceeded = false;
        int read;
        while ((read = await ReadAsync(reader, buffer, token).ConfigureAwait(false)) != 0) {
            int keep = Math.Min(read, 32768 - text.Length);
            text.Append(buffer, 0, keep);
            exceeded |= keep < read;
        }
        if (exceeded) throw new IOException("The print command exceeded its output limit.");
        return text.ToString();
    }
    private static async Task WaitForExitAsync(Process process, CancellationToken token) {
#if NET8_0_OR_GREATER
        await process.WaitForExitAsync(token).ConfigureAwait(false);
#else
        while (!process.HasExited) await Task.Delay(25, token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
#endif
    }

    private static async Task<int> ReadAsync(StreamReader reader, char[] buffer, CancellationToken token) {
#if NET8_0_OR_GREATER
        return await reader.ReadAsync(buffer.AsMemory(), token).ConfigureAwait(false);
#else
        token.ThrowIfCancellationRequested();
        int count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        return count;
#endif
    }

#if !NET8_0_OR_GREATER
    private static string QuoteProcessArgument(string value) {
        var result = new StringBuilder("\"");
        int slashes = 0;
        foreach (char character in value) {
            if (character == '\\') { slashes++; continue; }
            result.Append('\\', character == '"' ? slashes * 2 + 1 : slashes);
            result.Append(character);
            slashes = 0;
        }
        result.Append('\\', slashes * 2);
        return result.Append('"').ToString();
    }
#endif
}
