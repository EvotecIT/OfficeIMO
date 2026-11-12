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

    public static IppPrinterEndpoint ForUri(string printerName, Uri printerUri) => new(printerName, printerUri, null, null);

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

    internal static async Task<PdfPrinterCapabilities> GetCapabilitiesAsync(IppPrinterEndpoint endpoint, CancellationToken token) {
        using IppPrintClient client = endpoint.CreateClient();
        return MapCapabilities(endpoint.PrinterName, await client.GetCapabilitiesAsync(token).ConfigureAwait(false));
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

    internal static async Task<PdfPrintSubmission> SubmitAsync(IppPrinterEndpoint endpoint, PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        using IppPrintClient client = endpoint.CreateClient();
        IppPrinterCapabilities printer = await client.GetCapabilitiesAsync(token).ConfigureAwait(false);
        return await SubmitAsync(endpoint, client, printer, job, options, token).ConfigureAwait(false);
    }

    internal static async Task<PdfPrintSubmission> SubmitAsync(IppPrinterEndpoint endpoint, IppPrintClient client, IppPrinterCapabilities printer,
        PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        if (!(options.OutputFilePath == null || string.IsNullOrWhiteSpace(options.OutputFilePath)))
            throw new NotSupportedException("IPP printers do not write local output files.");
        IppDeliveryPlan plan = Plan(job.Plan, job.ContentKind, job.MaximumPixelsPerImage, options, printer, endpoint.PrinterName);
        using Stream document = plan.Kind switch {
            IppDocumentKind.VectorPdf => job.OpenContent(),
            IppDocumentKind.RasterPdf => job.ContentKind == PdfPrintContentKind.Raster && job.Plan.ColorMode != PdfPrintColorMode.Grayscale
                ? job.OpenContent() : new MemoryStream(job.RasterizeToPdf(options.MaximumRasterDpi, token), writable: false),
            _ => new PwgRasterJobStream(job, plan, token)
        };
        IppSubmission submission = await client.PrintAsync(document, plan.DocumentFormat, plan.JobOptions, token).ConfigureAwait(false);
        return new PdfPrintSubmission(endpoint.PrinterName, endpoint.FormatJobId(submission.Job.Id), job.Plan.Sheets.Count, options.Copies, null) {
            IgnoredAttributes = plan.Ignored.Concat(submission.UnsupportedAttributes).Distinct(StringComparer.Ordinal).ToArray()
        };
    }

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

        var plan = new IppDeliveryPlan();
        if (contentKind == PdfPrintContentKind.Vector && !grayscale && acceptsPdf) plan.Kind = IppDocumentKind.VectorPdf;
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
            else {
                job.Sides = sides;
                plan.Duplex = options.Duplex != PdfPrintDuplex.SingleSided;
                plan.Tumble = options.Duplex == PdfPrintDuplex.ShortEdge;
            }
        }

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
            plan.ColorSpace = (grayscale && gray) || !rgb ? PwgRasterColorSpace.Gray8 : PwgRasterColorSpace.Rgb8;
            plan.SheetBack = printer.PwgRasterSheetBack;
        } else {
            // Sheets are composed at their paper size; the printer must not scale them again.
            if (printer.PrintScaling.Contains("none")) job.PrintScaling = "none";
            IppOrientation? orientation = sheets.Sheets.All(sheet => sheet.PaperSize.Width.ToPoints() > sheet.PaperSize.Height.ToPoints()) ? IppOrientation.Landscape
                : sheets.Sheets.All(sheet => sheet.PaperSize.Width.ToPoints() <= sheet.PaperSize.Height.ToPoints()) ? IppOrientation.Portrait : null;
            if (orientation is { } value && printer.Orientations.Contains(value)) job.Orientation = value;
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
    public int RasterDpi { get; set; }
    public PwgRasterColorSpace ColorSpace { get; set; }
    /// <summary>pwg-raster-document-sheet-back, which says how back sides must be oriented.</summary>
    public string? SheetBack { get; set; }
}
