using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

internal static partial class CupsPdfPrinter {
    internal static async Task<PdfPrinterCapabilities> GetCapabilitiesAsync(CupsScheduler scheduler, string printerName, CancellationToken token) {
        IppQueueRoute? route = await ResolveIppQueueAsync(scheduler, printerName, token).ConfigureAwait(false);
        if (route is not null) {
            using (route.Client) return IppPdfPrinter.MapCapabilities(printerName, route.Printer);
        }
        if (!(await GetPrintersAsync(scheduler, token).ConfigureAwait(false)).Any(printer => printer.Name == printerName))
            throw new ArgumentException("The selected printer queue is not available.", nameof(printerName));
        return await ReadCapabilitiesAsync(printerName, token).ConfigureAwait(false);
    }

    /// <summary>Reads a job from the scheduler over IPP. Jobs submitted through lp are tracked the same way.</summary>
    internal static Task<PdfPrintJobStatus> GetJobAsync(CupsScheduler scheduler, string printerName, string jobId, CancellationToken token) =>
        IppPdfPrinter.GetJobAsync(TrackingEndpoint(scheduler, printerName), jobId, token);

    internal static Task<PdfPrintJobStatus> CancelJobAsync(CupsScheduler scheduler, string printerName, string jobId, CancellationToken token) =>
        IppPdfPrinter.CancelJobAsync(TrackingEndpoint(scheduler, printerName), jobId, token);

    // The scheduler finds a job by its number; the queue path only scopes the request.
    private static IppPrinterEndpoint TrackingEndpoint(CupsScheduler scheduler, string printerName) => scheduler.IsAvailable
        ? scheduler.Queue(printerName, isClass: false)
        : throw new PlatformNotSupportedException("Job tracking needs the CUPS scheduler over IPP, which is not reachable on this host.");

    /// <summary>A queue the scheduler answered for over IPP, with the client that reached it.</summary>
    internal sealed record IppQueueRoute(IppPrinterEndpoint Endpoint, IppPrintClient Client, IppPrinterCapabilities Printer);

    /// <summary>
    /// Reads the queue's attributes from the scheduler over IPP. Returns null, so the command-line tools are used, when the scheduler
    /// cannot be reached over IPP or does not host the queue (a discovered destination is created by lp on first use).
    /// </summary>
    internal static async Task<IppQueueRoute?> ResolveIppQueueAsync(CupsScheduler scheduler, string queue, CancellationToken token) {
        if (!scheduler.IsAvailable) return null;
        foreach (bool isClass in new[] { false, true }) {
            IppPrinterEndpoint endpoint = scheduler.Queue(queue, isClass);
            IppPrintClient client = endpoint.CreateClient();
            try {
                return new IppQueueRoute(endpoint, client, await client.GetCapabilitiesAsync(token).ConfigureAwait(false));
            } catch (IppException error) when (error.StatusCode == NotFound) {
                client.Dispose();
            } catch (IppException error) when (error.StatusCode is null) {
                client.Dispose();
                return null;
            } catch {
                client.Dispose();
                throw;
            }
        }
        return null;
    }

    private const int NotFound = 0x0406;

    private static async Task<PdfPrinterCapabilities> ReadCapabilitiesAsync(string printerName, CancellationToken token) {
        var choices = await RunAsync("lpoptions", ["-p", printerName, "-l"], token).ConfigureAwait(false);
        if (choices.ExitCode != 0) throw new IOException(choices.Error.Trim());
        var destination = await RunAsync("lpoptions", ["-p", printerName], token).ConfigureAwait(false);
        if (destination.ExitCode != 0) throw new IOException(destination.Error.Trim());
        return ParseCapabilities(printerName, choices.Output, destination.Output);
    }

    /// <summary>Lists the scheduler's queues over IPP together with the destinations <c>lpstat -e</c> reports, which include discovered printers.</summary>
    internal static async Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CupsScheduler scheduler, CancellationToken token) {
        IReadOnlyList<(string Name, bool IsClass)> installed = [];
        string? defaultName = null;
        bool reached = false;
        if (scheduler.IsAvailable) {
            try {
                using IppPrintClient client = scheduler.CreateSchedulerClient();
                (installed, defaultName) = await client.GetCupsQueuesAsync(token).ConfigureAwait(false);
                reached = true;
            } catch (IppException) {
                // Not reachable over IPP, or not a CUPS scheduler; the command-line tools still list destinations.
            }
        }
        if (!File.Exists("/usr/bin/lpstat")) {
            if (!reached) throw new PlatformNotSupportedException("Neither the CUPS scheduler nor the CUPS command-line tools are available.");
            return MergePrinters(installed.Select(queue => queue.Name), defaultName);
        }
        var queues = await RunAsync("lpstat", ["-e"], token).ConfigureAwait(false);
        if (queues.ExitCode != 0 && !reached) throw new IOException(queues.Error.Trim());
        var current = await RunAsync("lpstat", ["-d"], token).ConfigureAwait(false);
        IReadOnlyList<PdfPrinterInfo> listed = queues.ExitCode == 0 ? ParsePrinters(queues.Output, current.Output) : [];
        return MergePrinters(installed.Select(queue => queue.Name).Concat(listed.Select(printer => printer.Name)),
            defaultName ?? listed.FirstOrDefault(printer => printer.IsDefault)?.Name);
    }

    private static IReadOnlyList<PdfPrinterInfo> MergePrinters(IEnumerable<string> names, string? defaultName) =>
        names.Distinct(StringComparer.Ordinal).Select(name => new PdfPrinterInfo(name, name == defaultName, false))
            .OrderByDescending(printer => printer.IsDefault).ThenBy(printer => printer.Name, StringComparer.Ordinal).ToArray();

    /// <summary>Reads <c>lpstat -e</c> and <c>lpstat -d</c>. macOS localizes the messages regardless of the locale variables, so the default is recognized by its queue name.</summary>
    internal static IReadOnlyList<PdfPrinterInfo> ParsePrinters(string queues, string defaultDestination) {
        string[] names = queues.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(name => name.Trim()).Where(name => name.Length != 0)
            .Distinct(StringComparer.Ordinal).ToArray();
        string current = defaultDestination.Trim();
        int separator = current.LastIndexOf(": ", StringComparison.Ordinal);
        string? defaultName = separator < 0 ? null : current.Substring(separator + 2).Trim();
        return names.Select(name => new PdfPrinterInfo(name, name == defaultName, false))
            .OrderByDescending(printer => printer.IsDefault).ThenBy(printer => printer.Name, StringComparer.Ordinal).ToArray();
    }

    /// <summary>Finds the job identifier <c>queue-number</c> in localized <c>lp</c> output.</summary>
    internal static string? ParseJobId(string output, string printerName) {
        Match receipt = Regex.Match(output, @"(?<![^\s(])" + Regex.Escape(printerName) + @"-\d+(?![^\s),.])",
            RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        return receipt.Success ? receipt.Value : null;
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(CupsScheduler scheduler, PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!(options.OutputFilePath == null || string.IsNullOrWhiteSpace(options.OutputFilePath))) throw new NotSupportedException("CUPS file output is controlled by the selected printer queue.");
        IppQueueRoute? route = await ResolveIppQueueAsync(scheduler, options.PrinterName, token).ConfigureAwait(false);
        if (route is not null) {
            using (route.Client)
                return await IppPdfPrinter.SubmitAsync(route.Endpoint, route.Client, route.Printer, job, options, token).ConfigureAwait(false);
        }
        return await SubmitWithLpAsync(scheduler, job, options, token).ConfigureAwait(false);
    }

    /// <summary>Delivers through <c>lp</c> when the scheduler cannot be reached over IPP or has not created the destination yet.</summary>
    private static async Task<PdfPrintSubmission> SubmitWithLpAsync(CupsScheduler scheduler, PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!(await GetPrintersAsync(scheduler, token).ConfigureAwait(false)).Any(printer => printer.Name == options.PrinterName))
            throw new ArgumentException("The selected printer queue is not available.", nameof(options));
        PdfPrinterCapabilities capabilities = await ReadCapabilitiesAsync(options.PrinterName, token).ConfigureAwait(false);
        if (options.PaperSourceId is not null && !capabilities.PaperSources.Any(source => source.Id == options.PaperSourceId))
            throw new ArgumentException("The selected paper source is no longer available on this printer.", nameof(options));
        // The vector content keeps source colors, so grayscale jobs send the same pixels the preview shows.
        using Stream content = job.Plan.ColorMode == PdfPrintColorMode.Grayscale
            ? new MemoryStream(job.RasterizeToPdf(options.MaximumRasterDpi, token), writable: false)
            : job.OpenContent();
        return await SubmitSpoolAsync(job.Plan, options, CreateArguments(job.Plan, options, capabilities.Media, capabilities.DefaultMedia), content, token, RunAsync).ConfigureAwait(false);
    }

    /// <summary>Builds lp options. One media keyword applies to the whole job, so mixed-size jobs leave media selection to each PDF page.</summary>
    internal static List<string> CreateArguments(PdfPrintPlan plan, PdfPrintDeliveryOptions options, IReadOnlyList<PdfPrinterMedia> media, PdfPrinterMedia? defaultMedia) {
        var arguments = new List<string> {
            "-d", options.PrinterName, "-n", options.Copies.ToString(CultureInfo.InvariantCulture), "-t", options.DocumentName,
            "-o", "print-scaling=none", "-o", "number-up=1", "-o", "Collate=True"
        };
        PdfPageSize first = plan.Sheets[0].PaperSize;
        if (plan.Sheets.All(sheet => PdfPrinterMediaNames.SameSize(sheet.PaperSize, first))) {
            string keyword = PdfPrinterMediaNames.Match(media, defaultMedia, first)?.Id ?? PdfPrinterMediaNames.TryGetKeyword(first) ?? CustomMedia(first);
            arguments.Add("-o"); arguments.Add("media=" + keyword);
            if (plan.Sheets.All(sheet => sheet.PaperSize.Width.ToPoints() > sheet.PaperSize.Height.ToPoints())) {
                arguments.Add("-o"); arguments.Add("orientation-requested=4");
            } else if (plan.Sheets.All(sheet => sheet.PaperSize.Width.ToPoints() <= sheet.PaperSize.Height.ToPoints())) {
                arguments.Add("-o"); arguments.Add("orientation-requested=3");
            }
        }
        if (options.Duplex != PdfPrintDuplex.PrinterDefault) {
            arguments.Add("-o");
            arguments.Add("sides=" + (options.Duplex == PdfPrintDuplex.SingleSided ? "one-sided" : options.Duplex == PdfPrintDuplex.LongEdge ? "two-sided-long-edge" : "two-sided-short-edge"));
        }
        if (options.PaperSourceId is not null) {
            arguments.Add("-o"); arguments.Add(options.PaperSourceId);
        }
        return arguments;
    }

    private static string CustomMedia(PdfPageSize paper) => string.Format(CultureInfo.InvariantCulture, "Custom.{0:0.###}x{1:0.###}pt",
        Math.Min(paper.Width.ToPoints(), paper.Height.ToPoints()), Math.Max(paper.Width.ToPoints(), paper.Height.ToPoints()));

    internal delegate Task<(int ExitCode, string Output, string Error)> CupsCommandRunner(
        string command, IReadOnlyList<string> arguments, CancellationToken token, Action? started);

    internal static async Task<PdfPrintSubmission> SubmitSpoolAsync(PdfPrintPlan plan, PdfPrintDeliveryOptions options,
        IReadOnlyList<string> lpOptions, Stream content, CancellationToken token, CupsCommandRunner run) {
        token.ThrowIfCancellationRequested();
        string directory = OfficeIMO.Core.Internal.OfficeTemporaryDirectory.Create("officeimo-print-");
        bool submissionUncertain = false;
        string? jobId = null;
        Exception? failure = null;
        PdfPrintSubmission? submission = null;
        try {
            string path = Path.Combine(directory, "job.pdf");
            using (FileStream spool = OfficeIMO.Core.Internal.OfficeTemporaryFile.CreateAtPath(path, 81920, FileOptions.Asynchronous)) {
                await content.CopyToAsync(spool, 81920, token).ConfigureAwait(false);
            }
            var arguments = new List<string>(lpOptions) { "--", path };
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
            jobId = ParseJobId(result.Output, options.PrinterName)
                ?? throw new IOException("The print command returned no recognizable job acknowledgement.");
            submission = new(options.PrinterName, jobId, plan.Sheets.Count, options.Copies, null);
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
