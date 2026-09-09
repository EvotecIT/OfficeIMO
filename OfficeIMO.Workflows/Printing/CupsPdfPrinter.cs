using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

internal static class CupsPdfPrinter {
    internal static async Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken token) {
        var queues = await RunAsync("lpstat", ["-e"], token).ConfigureAwait(false);
        if (queues.ExitCode != 0) throw new IOException(queues.Error.Trim());
        var current = await RunAsync("lpstat", ["-d"], token).ConfigureAwait(false);
        const string prefix = "system default destination: ";
        string? defaultName = current.Output.Trim().StartsWith(prefix, StringComparison.Ordinal)
            ? current.Output.Trim()[prefix.Length..] : null;
        return queues.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.Ordinal).Select(name => new PdfPrinterInfo(name, name == defaultName, false))
            .OrderByDescending(printer => printer.IsDefault).ThenBy(printer => printer.Name, StringComparer.Ordinal).ToArray();
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!string.IsNullOrWhiteSpace(options.OutputFilePath)) throw new NotSupportedException("CUPS file output is controlled by the selected printer queue.");
        if (!(await GetPrintersAsync(token).ConfigureAwait(false)).Any(printer => printer.Name == options.PrinterName))
            throw new ArgumentException("The selected printer queue is not available.", nameof(options));
        PageSize paper = document.Sheets[0].Plan.PaperSize;
        double width = Math.Min(paper.Width, paper.Height), height = Math.Max(paper.Width, paper.Height);
        if (document.Sheets.Any(sheet => Math.Abs(Math.Min(sheet.Plan.PaperSize.Width, sheet.Plan.PaperSize.Height) - width) > 1 ||
            Math.Abs(Math.Max(sheet.Plan.PaperSize.Width, sheet.Plan.PaperSize.Height) - height) > 1))
            throw new NotSupportedException("One CUPS job must use a single paper size.");
        var pages = new List<PdfDocument>();
        foreach (PdfRenderedPrintSheet sheet in document.Sheets) {
            token.ThrowIfCancellationRequested();
            pages.Add(PdfDocument.CreateFromImages([new PdfImageDocumentSource(sheet.Png)],
                new PdfImageDocumentOptions { FixedPageSize = sheet.Plan.PaperSize, Margin = 0, Fit = OfficeImageFit.Stretch }, token));
        }
        byte[] bytes = (pages.Count == 1 ? pages[0] : PdfDocument.Merge(pages, token)).ToBytes();
        if (bytes.LongLength > 256L * 1024 * 1024) throw new InvalidOperationException("The CUPS spool document exceeds its byte limit.");
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-print-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        bool started = false;
        string? jobId = null;
        Exception? failure = null;
        try {
            string path = Path.Combine(directory, "sheets.pdf");
            await File.WriteAllBytesAsync(path, bytes, token).ConfigureAwait(false);
            var arguments = new List<string> {
                "-d", options.PrinterName, "-n", options.Copies.ToString(CultureInfo.InvariantCulture), "-t", options.DocumentName,
                "-o", "media=" + Media(width, height), "-o", "print-scaling=none", "-o", "number-up=1", "-o", "Collate=True"
            };
            if (document.Sheets.All(sheet => sheet.Plan.PaperSize.Width > sheet.Plan.PaperSize.Height)) {
                arguments.Add("-o"); arguments.Add("orientation-requested=4");
            } else if (document.Sheets.All(sheet => sheet.Plan.PaperSize.Width <= sheet.Plan.PaperSize.Height)) {
                arguments.Add("-o"); arguments.Add("orientation-requested=3");
            }
            if (options.Duplex != PdfPrintDuplex.PrinterDefault) {
                arguments.Add("-o");
                arguments.Add("sides=" + (options.Duplex == PdfPrintDuplex.SingleSided ? "one-sided" : options.Duplex == PdfPrintDuplex.LongEdge ? "two-sided-long-edge" : "two-sided-short-edge"));
            }
            arguments.Add("--"); arguments.Add(path);
            var result = await RunAsync("lp", arguments, token, () => started = true).ConfigureAwait(false);
            if (result.ExitCode != 0) throw new IOException(result.Error.Trim());
            Match receipt = Regex.Match(result.Output, @"request id is (\S+-\d+)", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            if (!receipt.Success) throw new IOException("The print command returned no recognizable job acknowledgement.");
            jobId = receipt.Groups[1].Value;
            return new(options.PrinterName, jobId, document.Sheets.Count, options.Copies, null);
        } catch (Exception error) {
            failure = error;
            if (started) throw new PdfPrintDeliveryException(jobId, error);
            throw;
        } finally {
            try { Directory.Delete(directory, recursive: true); }
            catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException) {
                Exception detail = new IOException($"Could not remove print staging at '{directory}'.", cleanup);
                if (failure is not null) detail = new AggregateException(failure, detail);
                if (started) throw new PdfPrintDeliveryException(jobId, detail);
                throw detail;
            }
        }
    }

    private static string Media(double width, double height) {
        if (Math.Abs(width - PageSizes.A4.Width) < 1 && Math.Abs(height - PageSizes.A4.Height) < 1) return "A4";
        if (Math.Abs(width - PageSizes.A3.Width) < 1 && Math.Abs(height - PageSizes.A3.Height) < 1) return "A3";
        if (Math.Abs(width - PageSizes.Letter.Width) < 1 && Math.Abs(height - PageSizes.Letter.Height) < 1) return "Letter";
        if (Math.Abs(width - PageSizes.Legal.Width) < 1 && Math.Abs(height - PageSizes.Legal.Height) < 1) return "Legal";
        return string.Format(CultureInfo.InvariantCulture, "Custom.{0:0.###}x{1:0.###}pt", width, height);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunAsync(string command, IReadOnlyList<string> arguments, CancellationToken token, Action? started = null) {
        string executable = "/usr/bin/" + command;
        if (!File.Exists(executable)) throw new PlatformNotSupportedException("CUPS command-line printing tools are not installed.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        start.Environment["LC_ALL"] = "C";
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using var process = new Process { StartInfo = start };
        timeout.Token.ThrowIfCancellationRequested();
        if (!process.Start()) throw new IOException("The print command could not start.");
        started?.Invoke();
        Task<string> output = ReadBoundedAsync(process.StandardOutput);
        Task<string> error = ReadBoundedAsync(process.StandardError);
        try {
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
            return (process.ExitCode, await output.ConfigureAwait(false), await error.ConfigureAwait(false));
        } catch {
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
            await process.WaitForExitAsync(CancellationToken.None).ConfigureAwait(false);
            try { await Task.WhenAll(output, error).ConfigureAwait(false); } catch (IOException) { }
            throw;
        }
    }

    private static async Task<string> ReadBoundedAsync(StreamReader reader) {
        var text = new StringBuilder();
        char[] buffer = new char[4096];
        bool exceeded = false;
        int read;
        while ((read = await reader.ReadAsync(buffer).ConfigureAwait(false)) != 0) {
            int keep = Math.Min(read, 32768 - text.Length);
            text.Append(buffer, 0, keep);
            exceeded |= keep < read;
        }
        if (exceeded) throw new IOException("The print command exceeded its output limit.");
        return text.ToString();
    }
}
