using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace OfficeIMO.Workflows.Tests;

/// <summary>
/// Qualifies the IPP client against CUPS's independent IPP Everywhere reference printer, started on a
/// loopback port with a private spool directory. Skipped where ippeveprinter is not installed.
/// </summary>
public sealed class IppEvePrinterTests {
    private const string IppEvePrinterPath = ExternalTool.IppEvePrinter;
    private static readonly TimeSpan JobTimeout = TimeSpan.FromSeconds(60);
    private static readonly byte[] Pdf = Encoding.ASCII.GetBytes(
        "%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n" +
        "3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF\n");

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task ReadsTheReferencePrinterCapabilities() {
        using var printer = EvePrinter.Start("-2", "-f", "application/pdf,image/pwg-raster");
        using var client = new IppPrintClient(printer.Uri);

        IppPrinterCapabilities capabilities = await client.GetCapabilitiesAsync();

        Assert.Equal(IppPrinterState.Idle, capabilities.State);
        Assert.False(string.IsNullOrEmpty(capabilities.MakeAndModel));
        Assert.Contains("application/pdf", capabilities.DocumentFormats);
        Assert.Contains("image/pwg-raster", capabilities.DocumentFormats);
        Assert.Equal(new[] { "one-sided", "two-sided-long-edge", "two-sided-short-edge" }, capabilities.Sides);
        Assert.Contains("iso_a4_210x297mm", capabilities.Media);
        Assert.NotNull(capabilities.MediaDefault);
        Assert.Contains(capabilities.MediaSizes, media => media.Width == 21000 && media.Height == 29700);
        Assert.NotNull(capabilities.MediaColDefault);
        Assert.NotEmpty(capabilities.ColorModes);
        Assert.NotEmpty(capabilities.Resolutions);
        Assert.Contains(new IppResolution(300, 300, IppResolutionUnits.DotsPerInch), capabilities.PwgRasterResolutions);
        Assert.Contains("2.0", capabilities.IppVersions);
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task PrintsPdfFromBytesAndUnknownLengthStreamsAndSpoolsIdenticalBytes() {
        using var printer = EvePrinter.Start("-2", "-f", "application/pdf,image/pwg-raster", "-c", "/usr/bin/true");
        using var client = new IppPrintClient(printer.Uri, new IppPrintClientOptions { RequestingUserName = "officeimo" });
        var options = new IppJobOptions { JobName = "ipp-pdf", Sides = "two-sided-long-edge", Media = "iso_a4_210x297mm", Orientation = IppOrientation.Portrait };

        Assert.Empty(await client.ValidateJobAsync("application/pdf", options));
        // The reference printer processes one job at a time, so each submission waits for the previous job.
        var submissions = new List<IppSubmission>();
        foreach (Stream document in new Stream[] { new MemoryStream(Pdf), new UnknownLengthStream(Pdf) }) {
            IppSubmission submission = await client.PrintAsync(document, "application/pdf", options);
            Assert.Empty(submission.UnsupportedAttributes);
            IppJobInfo finished = await client.WaitForJobAsync(submission.Job.Id, JobTimeout, TimeSpan.FromMilliseconds(100));
            Assert.Equal(IppJobState.Completed, finished.State);
            Assert.Equal(Pdf, File.ReadAllBytes(printer.SpooledDocument(submission.Job.Id, ".pdf")));
            submissions.Add(submission);
        }
        IReadOnlyList<IppJobInfo> completed = await client.GetJobsAsync(IppWhichJobs.Completed);
        Assert.Equal(submissions.Select(s => s.Job.Id), completed.Where(job => job.Name == "ipp-pdf").Select(job => job.Id).OrderBy(id => id));
    }

    [ExternalToolFact(IppEvePrinterPath, ExternalTool.IppEvePcl)]
    public async Task PrintsPwgRasterThatTheReferenceConverterProcesses() {
        // ippevepcl parses each PWG raster page; a malformed stream makes the job abort instead of complete.
        using var printer = EvePrinter.Start("-f", "image/pwg-raster", "-c", ExternalTool.Path(ExternalTool.IppEvePcl));
        using var client = new IppPrintClient(printer.Uri);
        byte[] raster = PwgRasterWriterTests.Write(
            PwgRasterWriterTests.Pattern(PwgRasterColorSpace.Gray8, 2480 / 4, 3508 / 4, seed: 5),
            PwgRasterWriterTests.Pattern(PwgRasterColorSpace.Rgb8, 2480 / 4, 3508 / 4, seed: 6));

        IppSubmission submission = await client.PrintAsync(raster, "image/pwg-raster", new IppJobOptions { JobName = "ipp-raster" });
        IppJobInfo finished = await client.WaitForJobAsync(submission.Job.Id, JobTimeout, TimeSpan.FromMilliseconds(100));

        Assert.Equal(IppJobState.Completed, finished.State);
        Assert.Equal(raster, File.ReadAllBytes(printer.SpooledDocument(submission.Job.Id, ".pwg")));
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task ServiceTracksAndCancelsAJobThatIsStillProcessing() {
        // Without a print command the reference printer simulates several seconds of processing per job.
        using var printer = EvePrinter.Start("-f", "application/pdf");
        var service = new PdfPrinterService();
        string uri = printer.Uri.AbsoluteUri;

        PdfPrintSubmission receipt = await service.SubmitAsync(TwoPageJob(PdfPrintColorMode.Color), new PdfPrintDeliveryOptions { PrinterName = uri, DocumentName = "ipp-cancel" });
        PdfPrintJobStatus running = await service.GetJobAsync(uri, receipt.JobId);
        Assert.False(running.IsTerminal);
        Assert.Equal(receipt.JobId, running.JobId);
        await Assert.ThrowsAsync<TimeoutException>(() => service.WaitForJobAsync(uri, receipt.JobId, TimeSpan.FromMilliseconds(300), TimeSpan.FromMilliseconds(100)));

        PdfPrintJobStatus canceled = await service.CancelJobAsync(uri, receipt.JobId);
        canceled = canceled.IsTerminal ? canceled : await service.WaitForJobAsync(uri, receipt.JobId, JobTimeout, TimeSpan.FromMilliseconds(100));
        Assert.Equal(PdfPrintJobState.Canceled, canceled.State);
        Assert.Contains("job-canceled-by-user", canceled.StateReasons);
        // Canceling a finished job reports it instead of failing, and unknown jobs are reported as removed.
        Assert.Equal(PdfPrintJobState.Canceled, (await service.CancelJobAsync(uri, receipt.JobId)).State);
        Assert.Equal(PdfPrintJobState.Removed, (await service.GetJobAsync(uri, "999")).State);
        Assert.Equal(PdfPrintJobState.Removed, (await service.CancelJobAsync(uri, "999")).State);
        await Assert.ThrowsAsync<ArgumentException>(() => service.GetJobAsync(uri, "printer-12"));
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task UnsupportedFormatsAreRejectedBeforeAJobExists() {
        using var printer = EvePrinter.Start("-f", "application/pdf");
        using var client = new IppPrintClient(printer.Uri);

        IppException error = await Assert.ThrowsAsync<IppException>(() => client.PrintAsync(Pdf, "image/urf"));

        // RFC 8011 allows either status for an unsupported document-format value.
        Assert.Contains(error.StatusCode, new int?[] { 0x040A, 0x040B });
        Assert.Empty(await client.GetJobsAsync());
        IppException missing = await Assert.ThrowsAsync<IppException>(() => client.GetJobAsync(999));
        Assert.Equal(0x0406, missing.StatusCode);
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task ServiceReadsCapabilitiesAndSendsVectorPdfToAPrinterThatAcceptsIt() {
        using var printer = EvePrinter.Start("-2", "-s", "10,10", "-f", "application/pdf,image/pwg-raster", "-c", "/usr/bin/true");
        var service = new PdfPrinterService();
        string uri = printer.Uri.AbsoluteUri;

        PdfPrinterCapabilities capabilities = await service.GetCapabilitiesAsync(uri);
        Assert.Equal(uri, capabilities.PrinterName);
        Assert.Contains(capabilities.Media, media => media.Id == "iso_a4_210x297mm");
        Assert.NotNull(capabilities.DefaultMedia?.Size);
        Assert.True(capabilities.SupportsDuplex);
        Assert.True(capabilities.SupportsColor);
        Assert.Contains(capabilities.PaperSources, source => source.Id == "main");

        PdfPrintJob job = TwoPageJob(PdfPrintColorMode.Color);
        PdfPrintSubmission receipt = await service.SubmitAsync(job, new PdfPrintDeliveryOptions {
            PrinterName = uri, DocumentName = "service-pdf", Duplex = PdfPrintDuplex.LongEdge, Copies = 2, PaperSourceId = "main"
        });

        Assert.Equal(uri, receipt.PrinterName);
        Assert.Empty(receipt.IgnoredAttributes);
        Assert.Equal(PdfPrintJobState.Completed, (await WaitAsync(service, receipt)).State);
        using var content = new MemoryStream();
        using (Stream source = job.OpenContent()) source.CopyTo(content);
        Assert.Equal(content.ToArray(), File.ReadAllBytes(printer.SpooledDocument(Number(receipt), ".pdf")));
    }

    [ExternalToolFact(IppEvePrinterPath, ExternalTool.IppEvePcl)]
    public async Task ServiceSendsPwgRasterToAPrinterWithoutPdfAndForGrayscaleJobs() {
        // ippevepcl parses each PWG raster page, so the job completes only when the raster stream is well formed.
        string conf = Path.Combine(Path.GetTempPath(), "officeimo-pwg-only-" + Guid.NewGuid().ToString("N") + ".conf");
        // An attribute file replaces the legacy flags; this printer accepts only color PWG raster and wants rotated back sides.
        File.WriteAllText(conf, "ATTR mimeMediaType document-format-supported image/pwg-raster\n" +
            "ATTR keyword pwg-raster-document-type-supported sgray_8,srgb_8\n" +
            "ATTR resolution pwg-raster-document-resolution-supported 150dpi,300dpi\n" +
            "ATTR keyword print-color-mode-supported auto,color,monochrome\n" +
            "ATTR keyword sides-supported one-sided,two-sided-long-edge,two-sided-short-edge\n" +
            "ATTR keyword pwg-raster-document-sheet-back rotated\n");
        try {
            using var rasterOnly = EvePrinter.Start("-a", conf, "-c", ExternalTool.Path(ExternalTool.IppEvePcl));
            using var pdfPrinter = EvePrinter.Start("-f", "application/pdf,image/pwg-raster", "-c", ExternalTool.Path(ExternalTool.IppEvePcl));
            var service = new PdfPrinterService();
            foreach ((EvePrinter printer, PdfPrintColorMode color) in new[] { (rasterOnly, PdfPrintColorMode.Color), (pdfPrinter, PdfPrintColorMode.Grayscale) }) {
                PdfPrintSubmission receipt = await service.SubmitAsync(TwoPageJob(color), new PdfPrintDeliveryOptions {
                    PrinterName = printer.Uri.AbsoluteUri, DocumentName = "service-raster", Duplex = PdfPrintDuplex.LongEdge, MaximumRasterDpi = 150
                });
                Assert.Equal(PdfPrintJobState.Completed, (await WaitAsync(service, receipt)).State);
                byte[] spooled = File.ReadAllBytes(printer.SpooledDocument(Number(receipt), ".pwg"));
                Assert.Equal("RaS2", Encoding.ASCII.GetString(spooled, 0, 4));
            }
        } finally {
            File.Delete(conf);
        }
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task CupsQueuesAreDeliveredOverIppWithLpStyleJobIdentifiers() {
        // The reference printer stands in for a scheduler queue: the CUPS route resolves "officeimo-test" to its URI.
        using var printer = EvePrinter.Start("-f", "application/pdf,image/pwg-raster", "-c", "/usr/bin/true");
        var scheduler = new CupsScheduler(new Uri(printer.Uri, "/"), null, (_, isClass) => isClass ? new Uri(printer.Uri, "/missing") : printer.Uri);

        CupsPdfPrinter.IppQueueRoute route = await CupsPdfPrinter.ResolveIppQueueAsync(scheduler, "officeimo-test", default)
            ?? throw new InvalidOperationException("The queue did not resolve over IPP.");
        using (route.Client) Assert.Contains("application/pdf", route.Printer.DocumentFormats);
        PdfPrintSubmission receipt = await CupsPdfPrinter.SubmitAsync(scheduler, TwoPageJob(PdfPrintColorMode.Color),
            new PdfPrintDeliveryOptions { PrinterName = "officeimo-test", DocumentName = "cups-route" }, new PrinterCapabilityCache(), default);

        Assert.Equal("officeimo-test", receipt.PrinterName);
        Assert.Matches(@"^officeimo-test-\d+$", receipt.JobId);
        var tracked = new PdfPrinterService(scheduler);
        Assert.Equal(PdfPrintJobState.Completed, (await WaitAsync(tracked, receipt)).State);
        Assert.True(File.Exists(printer.SpooledDocument(Number(receipt), ".pdf")));
    }

    [Fact]
    public async Task AnUnreachableSchedulerLeavesQueuesToTheCommandLineTools() {
        var closed = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
        closed.Start();
        int port = ((IPEndPoint)closed.LocalEndpoint).Port;
        closed.Stop();
        var scheduler = new CupsScheduler(new Uri("ipp://localhost:" + port + "/"), null);

        Assert.Null(await CupsPdfPrinter.ResolveIppQueueAsync(scheduler, "officeimo-test", default));
        Assert.Null(await CupsPdfPrinter.ResolveIppQueueAsync(new CupsScheduler(null, null), "officeimo-test", default));
    }

    [ExternalToolFact("/private/var/run/cupsd")]
    public async Task TheLocalSchedulerAnswersOverItsDomainSocket() {
        // Read-only: lists installed queues and looks up a queue that does not exist.
        Assert.True(CupsScheduler.Local.IsAvailable);
        using IppPrintClient client = CupsScheduler.Local.CreateSchedulerClient();
        (IReadOnlyList<(string Name, bool IsClass)> queues, _) = await client.GetCupsQueuesAsync();
        Assert.DoesNotContain(queues, queue => queue.Name.Length == 0);
        Assert.Null(await CupsPdfPrinter.ResolveIppQueueAsync(CupsScheduler.Local, "officeimo-no-such-queue-" + Guid.NewGuid().ToString("N"), default));
    }

    [Theory]
    [InlineData("localhost:8631", "ipp://localhost:8631/")]
    [InlineData("127.0.0.2:8631", "ipp://127.0.0.2:8631/")]
    [InlineData("[::1]:8631", "ipp://[::1]:8631/")]
    [InlineData("[::1]", "ipp://[::1]:631/")]
    [InlineData("::1", "ipp://[::1]:631/")]
    [InlineData("localhost/version=1.1", "ipp://localhost:631/")]
    [InlineData(null, "ipp://localhost:631/")]
    public void LoopbackSchedulersKeepTheirConfiguredHostAndPort(string? server, string expected) {
        CupsScheduler scheduler = CupsScheduler.FromEnvironment(server, _ => false);
        Assert.True(scheduler.IsAvailable);
        Assert.Equal(new Uri(expected + "printers/q"), scheduler.Queue("q", isClass: false).PrinterUri);
        Assert.Equal(new Uri(expected + "printers/q"), scheduler.JobQueue("q").PrinterUri);
        Assert.Equal(new Uri("ipp://localhost/classes/a%20b"), CupsScheduler.FromEnvironment(null, path => path == "/run/cups/cups.sock").Queue("a b", isClass: true).PrinterUri);
    }

    [Theory]
    [InlineData("print.example.com", "ipp://print.example.com:631/")]
    [InlineData("print.example.com:8631", "ipp://print.example.com:8631/")]
    [InlineData("[2001:db8::1]:631", "ipp://[2001:db8::1]:631/")]
    public void ARemoteSchedulerSubmitsThroughLpAndTracksJobsOverIpp(string server, string expected) {
        // Only the addresses are built here; nothing is contacted.
        CupsScheduler scheduler = CupsScheduler.FromEnvironment(server, _ => false);
        Assert.False(scheduler.IsAvailable);
        Assert.Throws<PlatformNotSupportedException>(() => scheduler.Queue("q", isClass: false));
        IppPrinterEndpoint jobs = scheduler.JobQueue("office");
        Assert.Equal(new Uri(expected + "printers/office"), jobs.PrinterUri);
        Assert.Equal(12, jobs.ParseJobId("office-12"));
    }

    [Theory]
    [InlineData("localhost:0")]
    [InlineData("localhost:http")]
    [InlineData("[::1")]
    [InlineData("bad host")]
    public void UnusableCupsServerValuesLeaveJobsToTheCommandLineTools(string server) {
        CupsScheduler scheduler = CupsScheduler.FromEnvironment(server, _ => false);
        Assert.False(scheduler.IsAvailable);
        Assert.Throws<PlatformNotSupportedException>(() => scheduler.JobQueue("q"));
    }

    [ExternalToolFact(IppEvePrinterPath)]
    public async Task PrintsAWordDocumentByUriWithItsConversionReport() {
        using var printer = EvePrinter.Start("-f", "application/pdf,image/pwg-raster", "-c", "/usr/bin/true");
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-print-any-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string docx = Path.Combine(directory, "Letter.docx"), pdf = Path.Combine(directory, "Letter.pdf");
            using (OfficeIMO.Word.WordDocument document = OfficeIMO.Word.WordDocument.Create(docx)) {
                document.AddParagraph("Printed straight from Word");
                document.Save();
            }
            OfficeIMO.Pdf.PdfDocument.Create(builder => builder.Page(page => page.Size(OfficeIMO.Pdf.PdfPageSizes.A4))).Save(pdf);
            var service = new PdfPrinterService();

            OfficePrintResult word = await service.PrintAsync(new OfficePrintRequest {
                Layout = new PdfPrintPlanRequest { InputPath = docx, PagesPerSheet = 2 }, UsePrinterPaper = true,
                Delivery = new PdfPrintDeliveryOptions { PrinterName = printer.Uri.AbsoluteUri, DocumentName = "Letter.docx" }
            });
            OfficePrintResult direct = await service.PrintAsync(new OfficePrintRequest {
                Layout = new PdfPrintPlanRequest { InputPath = pdf },
                Delivery = new PdfPrintDeliveryOptions { PrinterName = printer.Uri.AbsoluteUri, DocumentName = "Letter.pdf" }
            });

            Assert.NotNull(word.Conversion);
            Assert.Equal(PdfPrintContentKind.Vector, word.ContentKind);
            Assert.Null(direct.Conversion);
            foreach (OfficePrintResult result in new[] { word, direct }) {
                Assert.Equal(PdfPrintJobState.Completed, (await WaitAsync(service, result.Submission)).State);
                Assert.StartsWith("%PDF", Encoding.ASCII.GetString(File.ReadAllBytes(printer.SpooledDocument(Number(result.Submission), ".pdf")), 0, 4));
            }
            // The printer's default media is US Letter, so the converted document was laid out on it.
            OfficeIMO.Pdf.PdfPageLayoutInfo sheet = OfficeIMO.Pdf.PdfDocument.Load(printer.SpooledDocument(Number(word.Submission), ".pdf")).GetPrintablePageLayouts()[0];
            Assert.Equal((612, 792), (Math.Round(Math.Min(sheet.Width, sheet.Height)), Math.Round(Math.Max(sheet.Width, sheet.Height))));
            await Assert.ThrowsAsync<ArgumentException>(() => service.PrintAsync(new OfficePrintRequest {
                Layout = new PdfPrintPlanRequest { InputPath = docx, PdfPassword = "secret" },
                Delivery = new PdfPrintDeliveryOptions { PrinterName = printer.Uri.AbsoluteUri }
            }));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static PdfPrintJob TwoPageJob(PdfPrintColorMode color) =>
        PdfPrintJob.Create(OfficeIMO.Pdf.PdfDocument.Create(document => {
            document.Page(page => page.Size(OfficeIMO.Pdf.PdfPageSizes.A4).Background(OfficeIMO.OfficeColor.Red));
            document.Page(page => page.Size(OfficeIMO.Pdf.PdfPageSizes.A4).Background(OfficeIMO.OfficeColor.Blue));
        }), new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = OfficeIMO.Pdf.PdfPageSizes.A4, ColorMode = color });

    private static Task<PdfPrintJobStatus> WaitAsync(IPdfPrinterService service, PdfPrintSubmission receipt) =>
        service.WaitForJobAsync(receipt.PrinterName, receipt.JobId, JobTimeout, TimeSpan.FromMilliseconds(100));

    private static int Number(PdfPrintSubmission receipt) => int.Parse(receipt.JobId.Substring(receipt.JobId.LastIndexOf('-') + 1));

    private sealed class EvePrinter : IDisposable {
        private readonly Process _process;
        private readonly string _spool;

        private EvePrinter(Process process, string spool, int port) {
            _process = process;
            _spool = spool;
            Uri = new Uri($"ipp://localhost:{port}/ipp/print");
        }

        public Uri Uri { get; }

        public static EvePrinter Start(params string[] options) {
            string spool = Path.Combine(Path.GetTempPath(), "officeimo-ippeve-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(spool);
            int port = FreePort();
            var start = new ProcessStartInfo(ExternalTool.Path(IppEvePrinterPath)) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (string argument in new[] { "-p", port.ToString(), "-n", "localhost", "-r", "off", "-d", spool, "-k" }.Concat(options)) start.ArgumentList.Add(argument);
            start.ArgumentList.Add("OfficeIMO Test " + port);
            Process process = Process.Start(start)!;
            process.OutputDataReceived += (_, _) => { };
            process.ErrorDataReceived += (_, _) => { };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            var printer = new EvePrinter(process, spool, port);
            var clock = Stopwatch.StartNew();
            while (true) {
                try {
                    using var probe = new TcpClient();
                    probe.Connect("localhost", port);
                    return printer;
                } catch (SocketException) when (clock.Elapsed < TimeSpan.FromSeconds(15) && !process.HasExited) {
                    Thread.Sleep(100);
                } catch {
                    printer.Dispose();
                    throw;
                }
            }
        }

        public string SpooledDocument(int jobId, string extension) =>
            Directory.GetFiles(_spool, jobId + "-*" + extension).Single();

        public void Dispose() {
            try {
                if (!_process.HasExited) _process.Kill();
                _process.WaitForExit(5000);
            } catch (InvalidOperationException) {
            }
            _process.Dispose();
            try { Directory.Delete(_spool, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        private static int FreePort() {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
    }

    private sealed class UnknownLengthStream : MemoryStream {
        public UnknownLengthStream(byte[] data) : base(data) { }
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => base.Position; set => throw new NotSupportedException(); }
    }
}
