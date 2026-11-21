namespace OfficeIMO.Workflows;

/// <summary>IPP job states (RFC 8011 section 5.3.7).</summary>
internal enum IppJobState {
    /// <summary>Waiting to be processed.</summary>
    Pending = 3,
    /// <summary>Held and not yet eligible for processing.</summary>
    Held = 4,
    /// <summary>Being processed or printed.</summary>
    Processing = 5,
    /// <summary>Processing stopped, usually because the printer needs attention.</summary>
    Stopped = 6,
    /// <summary>Canceled by a user or operator. Terminal.</summary>
    Canceled = 7,
    /// <summary>Aborted by the printer. Terminal.</summary>
    Aborted = 8,
    /// <summary>Finished. Terminal; the printer considers the job delivered.</summary>
    Completed = 9
}

/// <summary>IPP printer states (RFC 8011 section 5.4.11).</summary>
internal enum IppPrinterState {
    /// <summary>Ready for new jobs.</summary>
    Idle = 3,
    /// <summary>Processing a job.</summary>
    Processing = 4,
    /// <summary>Stopped; see the state reasons.</summary>
    Stopped = 5
}

/// <summary>Requested page orientation (RFC 8011 section 5.2.10).</summary>
internal enum IppOrientation {
    /// <summary>Portrait.</summary>
    Portrait = 3,
    /// <summary>Landscape, rotated 90 degrees counter-clockwise.</summary>
    Landscape = 4,
    /// <summary>Landscape, rotated 90 degrees clockwise.</summary>
    ReverseLandscape = 5,
    /// <summary>Portrait, rotated 180 degrees.</summary>
    ReversePortrait = 6
}

/// <summary>Units of an <see cref="IppResolution"/>.</summary>
internal enum IppResolutionUnits {
    /// <summary>Dots per inch.</summary>
    DotsPerInch = 3,
    /// <summary>Dots per centimetre.</summary>
    DotsPerCentimeter = 4
}

/// <summary>Which jobs Get-Jobs returns.</summary>
internal enum IppWhichJobs {
    /// <summary>Jobs that have not reached a terminal state.</summary>
    NotCompleted,
    /// <summary>Jobs in a terminal state that the printer still retains.</summary>
    Completed
}

/// <summary>A printer resolution.</summary>
/// <param name="CrossFeed">Cross-feed (horizontal) resolution.</param>
/// <param name="Feed">Feed (vertical) resolution.</param>
/// <param name="Units">Resolution units.</param>
internal sealed record IppResolution(int CrossFeed, int Feed, IppResolutionUnits Units) {
    /// <inheritdoc />
    public override string ToString() => CrossFeed == Feed
        ? CrossFeed + (Units == IppResolutionUnits.DotsPerInch ? "dpi" : "dpcm")
        : CrossFeed + "x" + Feed + (Units == IppResolutionUnits.DotsPerInch ? "dpi" : "dpcm");
}

/// <summary>An inclusive range of document page numbers, starting at 1.</summary>
internal sealed record IppPageRange(int First, int Last);

/// <summary>A media size with optional source, type and margins, as in an IPP media-col collection.</summary>
/// <param name="Width">Width (x-dimension) in hundredths of a millimetre.</param>
/// <param name="Height">Height (y-dimension) in hundredths of a millimetre.</param>
internal sealed record IppMedia(int Width, int Height) {
    /// <summary>PWG self-describing size name, such as iso_a4_210x297mm, when known.</summary>
    public string? SizeName { get; init; }
    /// <summary>Input tray keyword, such as main or by-pass-tray.</summary>
    public string? Source { get; init; }
    /// <summary>Media type keyword, such as stationery or envelope.</summary>
    public string? Type { get; init; }
    /// <summary>Left margin in hundredths of a millimetre; 0 requests borderless output.</summary>
    public int? LeftMargin { get; init; }
    /// <summary>Top margin in hundredths of a millimetre.</summary>
    public int? TopMargin { get; init; }
    /// <summary>Right margin in hundredths of a millimetre.</summary>
    public int? RightMargin { get; init; }
    /// <summary>Bottom margin in hundredths of a millimetre.</summary>
    public int? BottomMargin { get; init; }
}

/// <summary>Printer capabilities and state from Get-Printer-Attributes. Missing attributes yield empty lists or null.</summary>
internal sealed class IppPrinterCapabilities {
    /// <summary>printer-make-and-model.</summary>
    public string? MakeAndModel { get; init; }
    /// <summary>printer-state.</summary>
    public IppPrinterState State { get; init; } = IppPrinterState.Idle;
    /// <summary>printer-state-reasons, such as media-empty-error, excluding "none".</summary>
    public IReadOnlyList<string> StateReasons { get; init; } = Array.Empty<string>();
    /// <summary>printer-state-message.</summary>
    public string? StateMessage { get; init; }
    /// <summary>document-format-supported MIME types, such as application/pdf or image/pwg-raster.</summary>
    public IReadOnlyList<string> DocumentFormats { get; init; } = Array.Empty<string>();
    /// <summary>media-supported keywords.</summary>
    public IReadOnlyList<string> Media { get; init; } = Array.Empty<string>();
    /// <summary>media-default keyword.</summary>
    public string? MediaDefault { get; init; }
    /// <summary>Fixed sizes listed in media-col-database (custom size ranges are omitted).</summary>
    public IReadOnlyList<IppMedia> MediaSizes { get; init; } = Array.Empty<IppMedia>();
    /// <summary>media-col-default.</summary>
    public IppMedia? MediaColDefault { get; init; }
    /// <summary>media-source-supported keywords, such as main or by-pass-tray.</summary>
    public IReadOnlyList<string> MediaSources { get; init; } = Array.Empty<string>();
    /// <summary>sides-supported keywords: one-sided, two-sided-long-edge, two-sided-short-edge.</summary>
    public IReadOnlyList<string> Sides { get; init; } = Array.Empty<string>();
    /// <summary>sides-default.</summary>
    public string? SidesDefault { get; init; }
    /// <summary>print-color-mode-supported keywords, such as color and monochrome.</summary>
    public IReadOnlyList<string> ColorModes { get; init; } = Array.Empty<string>();
    /// <summary>print-color-mode-default.</summary>
    public string? ColorModeDefault { get; init; }
    /// <summary>printer-resolution-supported.</summary>
    public IReadOnlyList<IppResolution> Resolutions { get; init; } = Array.Empty<IppResolution>();
    /// <summary>printer-resolution-default.</summary>
    public IppResolution? ResolutionDefault { get; init; }
    /// <summary>Upper bound of copies-supported; 1 when the printer does not make copies itself.</summary>
    public int MaxCopies { get; init; } = 1;
    /// <summary>orientation-requested-supported.</summary>
    public IReadOnlyList<IppOrientation> Orientations { get; init; } = Array.Empty<IppOrientation>();
    /// <summary>print-scaling-supported keywords, such as auto, fit, fill and none.</summary>
    public IReadOnlyList<string> PrintScaling { get; init; } = Array.Empty<string>();
    /// <summary>page-ranges-supported.</summary>
    public bool PageRangesSupported { get; init; }
    /// <summary>pwg-raster-document-resolution-supported, for image/pwg-raster jobs.</summary>
    public IReadOnlyList<IppResolution> PwgRasterResolutions { get; init; } = Array.Empty<IppResolution>();
    /// <summary>pwg-raster-document-type-supported, such as sgray_8 and srgb_8.</summary>
    public IReadOnlyList<string> PwgRasterTypes { get; init; } = Array.Empty<string>();
    /// <summary>pwg-raster-document-sheet-back: how back sides of two-sided PWG raster jobs must be oriented (normal, flipped, rotated or manual-tumble).</summary>
    public string? PwgRasterSheetBack { get; init; }
    /// <summary>ipp-versions-supported.</summary>
    public IReadOnlyList<string> IppVersions { get; init; } = Array.Empty<string>();
}

/// <summary>Job attributes for Print-Job and Validate-Job. Unset values keep the printer defaults.</summary>
internal sealed class IppJobOptions {
    /// <summary>job-name shown in the printer queue.</summary>
    public string JobName { get; set; } = "OfficeIMO document";
    /// <summary>copies, from 1 through 9999.</summary>
    public int? Copies { get; set; }
    /// <summary>sides keyword: one-sided, two-sided-long-edge or two-sided-short-edge.</summary>
    public string? Sides { get; set; }
    /// <summary>media keyword such as iso_a4_210x297mm. Cannot be combined with <see cref="MediaCol"/>.</summary>
    public string? Media { get; set; }
    /// <summary>media-col size, source, type and margins. Cannot be combined with <see cref="Media"/>.</summary>
    public IppMedia? MediaCol { get; set; }
    /// <summary>print-color-mode keyword, such as color or monochrome.</summary>
    public string? ColorMode { get; set; }
    /// <summary>orientation-requested.</summary>
    public IppOrientation? Orientation { get; set; }
    /// <summary>print-scaling keyword: auto, auto-fit, fill, fit or none.</summary>
    public string? PrintScaling { get; set; }
    /// <summary>page-ranges; send only when the printer reports page-ranges-supported.</summary>
    public IReadOnlyList<IppPageRange>? PageRanges { get; set; }
    /// <summary>printer-resolution.</summary>
    public IppResolution? Resolution { get; set; }
    /// <summary>ipp-attribute-fidelity: true makes the printer reject the job instead of ignoring unsupported attributes.</summary>
    public bool AttributeFidelity { get; set; }

    internal IppJobOptions Snapshot() {
        var copy = (IppJobOptions)MemberwiseClone();
        if (copy.JobName is null || string.IsNullOrWhiteSpace(copy.JobName) || copy.JobName.Any(char.IsControl))
            throw new ArgumentException("A non-empty job name without control characters is required.", nameof(JobName));
        if (copy.JobName.Length > 255) throw new ArgumentException("The job name cannot exceed 255 characters.", nameof(JobName));
        if (copy.Copies is < 1 or > 9999) throw new ArgumentOutOfRangeException(nameof(Copies));
        Keyword(copy.Sides, nameof(Sides));
        Keyword(copy.Media, nameof(Media));
        Keyword(copy.ColorMode, nameof(ColorMode));
        Keyword(copy.PrintScaling, nameof(PrintScaling));
        if (copy.Media is not null && copy.MediaCol is not null) throw new ArgumentException("Set either Media or MediaCol, not both.", nameof(MediaCol));
        if (copy.MediaCol is { } media) {
            if (media.Width <= 0 || media.Height <= 0) throw new ArgumentOutOfRangeException(nameof(MediaCol), "Media dimensions must be positive.");
            if (media.LeftMargin < 0 || media.TopMargin < 0 || media.RightMargin < 0 || media.BottomMargin < 0)
                throw new ArgumentOutOfRangeException(nameof(MediaCol), "Margins cannot be negative.");
            Keyword(media.SizeName, nameof(MediaCol));
            Keyword(media.Source, nameof(MediaCol));
            Keyword(media.Type, nameof(MediaCol));
        }
        if (copy.Orientation is { } orientation && !OfficeRuntimeCompatibility.IsDefined(orientation)) throw new ArgumentOutOfRangeException(nameof(Orientation));
        if (copy.PageRanges is not null) {
            var ranges = copy.PageRanges.ToArray();
            if (ranges.Length is 0 or > 256) throw new ArgumentException("Page ranges must contain 1 to 256 entries.", nameof(PageRanges));
            int previous = 0;
            foreach (IppPageRange? range in ranges) {
                if (range is null || range.First < 1 || range.Last < range.First || range.First <= previous)
                    throw new ArgumentException("Page ranges must be ascending, non-overlapping and start at 1.", nameof(PageRanges));
                previous = range.Last;
            }
            copy.PageRanges = ranges;
        }
        if (copy.Resolution is { } resolution && (resolution.CrossFeed <= 0 || resolution.Feed <= 0 || !OfficeRuntimeCompatibility.IsDefined(resolution.Units)))
            throw new ArgumentOutOfRangeException(nameof(Resolution));
        return copy;
    }

    internal static void Keyword(string? value, string parameter) {
        if (value is null) return;
        if (value.Length is 0 or > 255) throw new ArgumentException("IPP keywords must contain 1 to 255 characters.", parameter);
        foreach (char c in value) {
            if (!(c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9' or '-' or '_' or '.'))
                throw new ArgumentException($"'{value}' is not a valid IPP keyword.", parameter);
        }
    }
}

/// <summary>A job's state as reported by the printer.</summary>
/// <param name="Id">job-id, unique within the printer.</param>
/// <param name="State">job-state.</param>
internal sealed record IppJobInfo(int Id, IppJobState State) {
    /// <summary>job-uri.</summary>
    public string? Uri { get; init; }
    /// <summary>job-state-reasons, such as job-printing or job-canceled-by-user, excluding "none".</summary>
    public IReadOnlyList<string> StateReasons { get; init; } = Array.Empty<string>();
    /// <summary>job-state-message.</summary>
    public string? StateMessage { get; init; }
    /// <summary>job-name.</summary>
    public string? Name { get; init; }
    /// <summary>True for canceled, aborted and completed jobs, whose state no longer changes.</summary>
    public bool IsTerminal => State is IppJobState.Canceled or IppJobState.Aborted or IppJobState.Completed;
}

/// <summary>Printer acceptance of a Print-Job request.</summary>
/// <remarks>Acceptance is not completion; use <see cref="IppPrintClient.WaitForJobAsync"/> to follow the job to a terminal state.</remarks>
/// <param name="Job">The created job.</param>
/// <param name="UnsupportedAttributes">Requested attributes the printer ignored or substituted.</param>
internal sealed record IppSubmission(IppJobInfo Job, IReadOnlyList<string> UnsupportedAttributes);

/// <summary>Connection settings for <see cref="IppPrintClient"/>.</summary>
internal sealed class IppPrintClientOptions {
    /// <summary>requesting-user-name sent with each request, or null to omit it.</summary>
    public string? RequestingUserName { get; set; }
    /// <summary>Limit for each request, including the document upload and response. Defaults to five minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);
    /// <summary>Largest response accepted from the printer, in bytes. Defaults to 8 MiB.</summary>
    public int MaxResponseBytes { get; set; } = 8 * 1024 * 1024;

    internal IppPrintClientOptions Snapshot() {
        if (Timeout <= TimeSpan.Zero && Timeout != System.Threading.Timeout.InfiniteTimeSpan) throw new ArgumentOutOfRangeException(nameof(Timeout));
        if (MaxResponseBytes is < 1024 or > 256 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxResponseBytes));
        if (RequestingUserName is not null && (RequestingUserName.Length is 0 or > 255 || RequestingUserName.Any(char.IsControl)))
            throw new ArgumentException("The requesting user name is invalid.", nameof(RequestingUserName));
        return (IppPrintClientOptions)MemberwiseClone();
    }
}
