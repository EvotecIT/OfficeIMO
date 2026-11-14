using System.Globalization;
using System.Net.Http;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>An IPP printer and how to reach it: a printer URI, or a CUPS queue on the local scheduler.</summary>
internal sealed class IppPrinterEndpoint {
    private readonly Func<HttpMessageHandler>? _createHandler;
    private readonly string? _jobPrefix;

    private IppPrinterEndpoint(string printerName, Uri printerUri, Func<HttpMessageHandler>? createHandler, string? jobPrefix) {
        PrinterName = printerName;
        PrinterUri = printerUri;
        _createHandler = createHandler;
        _jobPrefix = jobPrefix;
    }

    /// <summary>The name callers use: the URI for a printer addressed by URI, or the queue name.</summary>
    public string PrinterName { get; }
    public Uri PrinterUri { get; }

    public static IppPrinterEndpoint ForUri(string printerName, Uri printerUri, Func<HttpMessageHandler>? createHandler = null) =>
        new(printerName, printerUri, createHandler, null);

    /// <summary>A CUPS queue; job identifiers take the lp form <c>queue-number</c>.</summary>
    public static IppPrinterEndpoint ForCupsQueue(string queue, Uri printerUri, Func<HttpMessageHandler>? createHandler) =>
        new(queue, printerUri, createHandler, queue + "-");

    public IppPrintClient CreateClient() => new(PrinterUri, _createHandler?.Invoke() ?? IppHttpTransport.CreateDefaultHandler(), true,
        new IppPrintClientOptions { RequestingUserName = RequestingUserName() });

    public string FormatJobId(int id) => _jobPrefix + id.ToString(CultureInfo.InvariantCulture);

    public int ParseJobId(string jobId) {
        string number = _jobPrefix is not null && jobId.StartsWith(_jobPrefix, StringComparison.Ordinal) ? jobId.Substring(_jobPrefix.Length) : jobId;
        if (!int.TryParse(number, NumberStyles.None, CultureInfo.InvariantCulture, out int id) || id < 1)
            throw new ArgumentException($"'{jobId}' is not a job identifier of {PrinterName}.", nameof(jobId));
        return id;
    }

    /// <summary>The operating-system user; CUPS lets job owners cancel their own jobs.</summary>
    private static string? RequestingUserName() {
        string name;
        try { name = Environment.UserName; } catch (PlatformNotSupportedException) { return null; }
        return name is { Length: > 0 and <= 255 } && !name.Any(char.IsControl) ? name : null;
    }
}

/// <summary>Delivers print jobs to IPP printers: vector PDF where accepted, PWG raster otherwise.</summary>
internal static class IppPdfPrinter {
    internal const string PdfFormat = "application/pdf", PwgRasterFormat = "image/pwg-raster";

    internal static async Task<PdfPrinterCapabilities> GetCapabilitiesAsync(IppPrinterEndpoint endpoint, PrinterCapabilityCache cache, CancellationToken token) {
        using IppPrintClient client = endpoint.CreateClient();
        IppPrinterCapabilities printer = await client.GetCapabilitiesAsync(token).ConfigureAwait(false);
        cache.Set(endpoint.PrinterName, printer);
        return MapCapabilities(endpoint.PrinterName, printer);
    }

    internal static PdfPrinterCapabilities MapCapabilities(string printerName, IppPrinterCapabilities printer) {
        var media = new List<PdfPrinterMedia>();
        foreach (string keyword in printer.Media.Count > 0 ? printer.Media : printer.MediaSizes.Select(size => size.SizeName).OfType<string>()) {
            // custom_min_* and custom_max_* describe a custom-size range, not a selectable size.
            if (keyword.StartsWith("custom_", StringComparison.OrdinalIgnoreCase) || media.Any(item => item.Id == keyword)) continue;
            IppMedia? described = printer.MediaSizes.FirstOrDefault(size => size.SizeName == keyword);
            media.Add(new PdfPrinterMedia(keyword, keyword, described is null ? PdfPrinterMediaNames.TryGetSize(keyword) : Size(described)));
        }
        string? defaultName = printer.MediaDefault ?? printer.MediaColDefault?.SizeName;
        PdfPrinterMedia? defaultMedia = media.FirstOrDefault(item => item.Id == defaultName);
        if (defaultMedia is null && printer.MediaColDefault is { } colDefault)
            defaultMedia = new PdfPrinterMedia(defaultName ?? "default", defaultName ?? "Printer default", Size(colDefault));
        return new PdfPrinterCapabilities {
            PrinterName = printerName,
            Media = media,
            DefaultMedia = defaultMedia,
            PaperSources = printer.MediaSources.Where(source => source != "auto").Select(source => new PdfPaperSourceInfo(source, source)).ToArray(),
            SupportsDuplex = printer.Sides.Count == 0 ? null : printer.Sides.Any(side => side.StartsWith("two-sided", StringComparison.Ordinal)),
            SupportsColor = printer.ColorModes.Count == 0 ? null : printer.ColorModes.Contains("color"),
            SupportsCollation = null,
            MaximumCopies = printer.MaxCopies,
            Resolutions = printer.Resolutions.Select(Dpi).Distinct().ToArray()
        };
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(IppPrinterEndpoint endpoint, PdfPrintJob job, PdfPrintDeliveryOptions options,
        PrinterCapabilityCache cache, CancellationToken token) {
        using IppPrintClient client = endpoint.CreateClient();
        if (!cache.TryGet(endpoint.PrinterName, out IppPrinterCapabilities printer)) {
            printer = await client.GetCapabilitiesAsync(token).ConfigureAwait(false);
            cache.Set(endpoint.PrinterName, printer);
        }
        return await SubmitAsync(endpoint, client, printer, job, options, token).ConfigureAwait(false);
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(IppPrinterEndpoint endpoint, IppPrintClient client, IppPrinterCapabilities printer,
        PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!(options.OutputFilePath == null || string.IsNullOrWhiteSpace(options.OutputFilePath)))
            throw new NotSupportedException("IPP printers do not write local output files.");
        IppDeliveryPlan plan = Plan(job.Plan, job.ContentKind, job.MaximumPixelsPerImage, options, printer, endpoint.PrinterName);
        // PDF documents are rewound and PWG raster is encoded again if a busy printer asks for the job to be sent later.
        Func<Stream> openDocument;
        if (plan.Kind == IppDocumentKind.VectorPdf || plan.Kind == IppDocumentKind.RasterPdf && job.ContentKind == PdfPrintContentKind.Raster && job.Plan.ColorMode != PdfPrintColorMode.Grayscale) {
            openDocument = job.OpenContent;
        } else if (plan.Kind == IppDocumentKind.RasterPdf) {
            byte[] rasterPdf = job.RasterizeToPdf(options.MaximumRasterDpi, token);
            openDocument = () => new MemoryStream(rasterPdf, writable: false);
        } else {
            openDocument = () => new PwgRasterJobStream(job, plan, token);
        }
        IppSubmission submission = await client.PrintAsync(openDocument, plan.DocumentFormat, plan.JobOptions, token).ConfigureAwait(false);
        return new PdfPrintSubmission(endpoint.PrinterName, endpoint.FormatJobId(submission.Job.Id), job.Plan.Sheets.Count, options.Copies, null) {
            IgnoredAttributes = plan.Ignored.Concat(submission.UnsupportedAttributes).Distinct(StringComparer.Ordinal).ToArray()
        };
    }

    internal static async Task<PdfPrintJobStatus> GetJobAsync(IppPrinterEndpoint endpoint, string jobId, CancellationToken token) {
        int id = endpoint.ParseJobId(jobId);
        using IppPrintClient client = endpoint.CreateClient();
        return await ReadJobAsync(endpoint, client, id, jobId, token).ConfigureAwait(false);
    }

    /// <summary>Sends Cancel-Job, then reports the job. A job that already finished or is unknown is reported, not treated as an error.</summary>
    internal static async Task<PdfPrintJobStatus> CancelJobAsync(IppPrinterEndpoint endpoint, string jobId, CancellationToken token) {
        int id = endpoint.ParseJobId(jobId);
        using IppPrintClient client = endpoint.CreateClient();
        try {
            await client.CancelJobAsync(id, token).ConfigureAwait(false);
        } catch (IppException error) when (error.StatusCode is StatusNotPossible or StatusGone or StatusNotFound) {
            PdfPrintJobStatus current = await ReadJobAsync(endpoint, client, id, jobId, token).ConfigureAwait(false);
            if (!current.IsTerminal) throw;
            return current;
        }
        return await ReadJobAsync(endpoint, client, id, jobId, token).ConfigureAwait(false);
    }

    private const int StatusNotPossible = 0x0404, StatusNotFound = 0x0406, StatusGone = 0x0407;

    private static async Task<PdfPrintJobStatus> ReadJobAsync(IppPrinterEndpoint endpoint, IppPrintClient client, int id, string jobId, CancellationToken token) {
        IppJobInfo job;
        try {
            job = await client.GetJobAsync(id, token).ConfigureAwait(false);
        } catch (IppException error) when (error.StatusCode is StatusNotFound or StatusGone) {
            return new PdfPrintJobStatus(endpoint.PrinterName, jobId, PdfPrintJobState.Removed);
        }
        return MapJob(endpoint.PrinterName, jobId, job);
    }

    internal static PdfPrintJobStatus MapJob(string printerName, string jobId, IppJobInfo job) =>
        new(printerName, jobId, job.State switch {
            IppJobState.Pending => PdfPrintJobState.Pending,
            IppJobState.Held => PdfPrintJobState.Held,
            IppJobState.Processing => PdfPrintJobState.Processing,
            IppJobState.Stopped => PdfPrintJobState.Stopped,
            IppJobState.Canceled => PdfPrintJobState.Canceled,
            IppJobState.Aborted => PdfPrintJobState.Aborted,
            _ => PdfPrintJobState.Completed
        }) { StateReasons = job.StateReasons, StateMessage = job.StateMessage };

    /// <summary>Chooses the document format and job attributes for a printer.</summary>
    internal static IppDeliveryPlan Plan(PdfPrintPlan sheets, PdfPrintContentKind contentKind, long maximumPixels, PdfPrintDeliveryOptions options,
        IppPrinterCapabilities printer, string printerName) {
        bool grayscale = sheets.ColorMode == PdfPrintColorMode.Grayscale;
        bool acceptsPdf = printer.DocumentFormats.Contains(PdfFormat, StringComparer.OrdinalIgnoreCase);
        int[] rasterResolutions = printer.DocumentFormats.Contains(PwgRasterFormat, StringComparer.OrdinalIgnoreCase)
            ? printer.PwgRasterResolutions.Select(Dpi).Where(dpi => dpi.HorizontalDpi == dpi.VerticalDpi).Select(dpi => dpi.HorizontalDpi).Distinct().OrderBy(dpi => dpi).ToArray()
            : [];
        bool rgb = printer.PwgRasterTypes.Contains("srgb_8"), gray = printer.PwgRasterTypes.Contains("sgray_8");
        bool acceptsPwg = rasterResolutions.Length > 0 && (rgb || gray);

        // A printer that converts PDF to monochrome itself is sent the vector PDF; otherwise grayscale sheets are rasterized as the preview shows them.
        bool printerConvertsGray = !options.RasterizeGrayscale && printer.ColorModes.Contains("monochrome");
        var plan = new IppDeliveryPlan();
        if (contentKind == PdfPrintContentKind.Vector && acceptsPdf && (!grayscale || printerConvertsGray)) plan.Kind = IppDocumentKind.VectorPdf;
        else if (acceptsPwg) plan.Kind = IppDocumentKind.PwgRaster;
        else if (acceptsPdf) plan.Kind = IppDocumentKind.RasterPdf;
        else throw new NotSupportedException("The printer accepts neither application/pdf nor a supported PWG raster type (sgray_8 or srgb_8). It lists: "
            + string.Join(", ", printer.DocumentFormats) + ".");
        plan.DocumentFormat = plan.Kind == IppDocumentKind.PwgRaster ? PwgRasterFormat : PdfFormat;

        PdfPrinterCapabilities mapped = MapCapabilities(printerName, printer);
        var job = new IppJobOptions { JobName = JobName(options.DocumentName) };
        plan.JobOptions = job;

        int maximumCopies = Math.Max(1, printer.MaxCopies);
        if (options.Copies > 1) {
            // PWG raster pages are not stored by every printer, so multi-sheet raster copies are sent as collated sheet sequences.
            bool printerCopies = options.Copies <= maximumCopies && (plan.Kind != IppDocumentKind.PwgRaster || sheets.Sheets.Count == 1);
            if (printerCopies) job.Copies = options.Copies;
            else if (plan.Kind == IppDocumentKind.PwgRaster) plan.SoftwareCopies = options.Copies;
            else throw new NotSupportedException($"The printer makes at most {maximumCopies} copies of a PDF document and does not accept PWG raster.");
        }

        if (options.Duplex != PdfPrintDuplex.PrinterDefault) {
            string sides = options.Duplex switch {
                PdfPrintDuplex.SingleSided => "one-sided", PdfPrintDuplex.LongEdge => "two-sided-long-edge", _ => "two-sided-short-edge"
            };
            if (printer.Sides.Count > 0 && !printer.Sides.Contains(sides)) plan.Ignored.Add("sides");
            else job.Sides = sides;
        }
        // A job without sides prints with the printer's sides-default, which decides where software copies and back sides fall.
        string? effectiveSides = job.Sides ?? printer.SidesDefault;
        plan.Duplex = effectiveSides is not null && effectiveSides.StartsWith("two-sided", StringComparison.Ordinal);
        plan.Tumble = effectiveSides == "two-sided-short-edge";

        PdfPageSize first = sheets.Sheets[0].PaperSize;
        bool uniform = sheets.Sheets.All(sheet => PdfPrinterMediaNames.SameSize(sheet.PaperSize, first));
        PdfPrinterMedia?[] matches = sheets.Sheets.Select(sheet => PdfPrinterMediaNames.Match(mapped.Media, mapped.DefaultMedia, sheet.PaperSize)).ToArray();
        plan.MediaNames = matches.Select(match => match?.Id).ToArray();
        if (options.PaperSourceId is not null) {
            if (!printer.MediaSources.Contains(options.PaperSourceId))
                throw new ArgumentException("The selected paper source is not available on this printer.", nameof(options));
            if (!uniform) throw new NotSupportedException("A paper source can be selected only for jobs that use one paper size.");
            job.MediaCol = MediaCol(first, matches[0]) with { Source = options.PaperSourceId };
        } else if (uniform) {
            // A size the printer does not list can be requested by dimensions only from printers that describe media as collections.
            bool acceptsMediaCol = printer.MediaSizes.Count > 0 || printer.MediaColDefault is not null || printer.MediaSources.Count > 0;
            if (plan.MediaNames[0] is { } keyword) job.Media = keyword;
            else if (acceptsMediaCol) job.MediaCol = MediaCol(first, null);
            else plan.Ignored.Add("media");
        }

        if (plan.Kind == IppDocumentKind.PwgRaster) {
            double largest = sheets.Sheets.Max(sheet => sheet.PaperSize.Width.ToPoints() / 72D * (sheet.PaperSize.Height.ToPoints() / 72D));
            double budget = Math.Floor(Math.Sqrt(maximumPixels / largest));
            int[] withinLimit = rasterResolutions.Where(dpi => dpi <= options.MaximumRasterDpi && dpi <= budget).ToArray();
            plan.RasterDpi = withinLimit.Length > 0 ? withinLimit[withinLimit.Length - 1]
                : rasterResolutions[0] <= budget ? rasterResolutions[0]
                : throw new InvalidOperationException("The printer's lowest PWG raster resolution exceeds the job's pixel budget.");
            // A printer that lists only resolutions above the caller's limit, such as many lasers and CUPS queues at 600 dpi only, gets sheets
            // rendered at the largest whole fraction of that resolution within the limit, each pixel repeated. Rendering four times the pixels
            // would cost four times the time without the detail the caller asked for.
            if (withinLimit.Length == 0) {
                int device = plan.RasterDpi;
                plan.RasterScale = Math.Max(1, Enumerable.Range(2, 7)
                    .FirstOrDefault(scale => device % scale == 0 && device / scale <= options.MaximumRasterDpi && device / scale >= 36));
            }
            plan.ColorSpace = (grayscale && gray) || !rgb ? PwgRasterColorSpace.Gray8 : PwgRasterColorSpace.Rgb8;
            plan.SheetBack = printer.PwgRasterSheetBack;
        } else {
            // Sheets are composed at their paper size and orientation; the printer must not scale or rotate them again.
            if (printer.PrintScaling.Contains("none")) job.PrintScaling = "none";
        }
        string colorMode = grayscale ? "monochrome" : "color";
        if (printer.ColorModes.Contains(colorMode)) job.ColorMode = colorMode;
        return plan;
    }

    /// <summary>A media-col for the paper, using the printer's own dimensions for a matched size so they compare equal to its media database.</summary>
    private static IppMedia MediaCol(PdfPageSize paper, PdfPrinterMedia? match) {
        if (match?.Size is { } size) paper = size;
        string? sizeName = match?.Id;
        double shortSide = Math.Min(paper.Width.ToPoints(), paper.Height.ToPoints()), longSide = Math.Max(paper.Width.ToPoints(), paper.Height.ToPoints());
        return new IppMedia(Hundredths(shortSide), Hundredths(longSide)) { SizeName = sizeName };
    }

    private static int Hundredths(double points) => (int)Math.Round(points / 72D * 2540D, MidpointRounding.AwayFromZero);

    private static PdfPageSize Size(IppMedia media) => PdfPrinterMediaNames.Portrait(media.Width / 2540D * 72D, media.Height / 2540D * 72D);

    private static PdfPrinterResolution Dpi(IppResolution resolution) => resolution.Units == IppResolutionUnits.DotsPerInch
        ? new(resolution.CrossFeed, resolution.Feed)
        : new((int)Math.Round(resolution.CrossFeed * 2.54), (int)Math.Round(resolution.Feed * 2.54));

    private static string JobName(string documentName) {
        string name = new(documentName.Select(character => char.IsControl(character) ? ' ' : character).ToArray());
        name = name.Trim();
        if (name.Length == 0) name = "OfficeIMO document";
        return name.Length > 255 ? name.Substring(0, 255) : name;
    }
}

internal enum IppDocumentKind {
    /// <summary>The job's vector PDF content.</summary>
    VectorPdf,
    /// <summary>A PDF of raster sheets, for printers without PWG raster support.</summary>
    RasterPdf,
    /// <summary>PWG raster rendered from the job's sheets.</summary>
    PwgRaster
}

/// <summary>Document format and job attributes chosen for one submission.</summary>
internal sealed class IppDeliveryPlan {
    public IppDocumentKind Kind { get; set; }
    public string DocumentFormat { get; set; } = IppPdfPrinter.PdfFormat;
    public IppJobOptions JobOptions { get; set; } = new();
    /// <summary>Requested attributes left out because the printer does not support them.</summary>
    public List<string> Ignored { get; } = new();
    /// <summary>Times the sheet sequence is repeated in the document when the printer does not make the copies.</summary>
    public int SoftwareCopies { get; set; } = 1;
    public bool Duplex { get; set; }
    public bool Tumble { get; set; }
    /// <summary>PWG media name for each sheet, or null where the printer lists no matching size.</summary>
    public IReadOnlyList<string?> MediaNames { get; set; } = [];
    /// <summary>Device resolution written to the PWG raster headers.</summary>
    public int RasterDpi { get; set; }
    /// <summary>Device pixels per rendered pixel along each edge; sheets render at <see cref="RasterDpi"/> divided by this.</summary>
    public int RasterScale { get; set; } = 1;
    public PwgRasterColorSpace ColorSpace { get; set; }
    /// <summary>pwg-raster-document-sheet-back, which says how back sides must be oriented.</summary>
    public string? SheetBack { get; set; }
}
