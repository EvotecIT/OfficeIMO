using System.Net.Http;

namespace OfficeIMO.Workflows;

/// <summary>
/// Prints to an IPP printer by URI (ipp:// or ipps://) without an installed queue, driver or CUPS tools.
/// </summary>
/// <remarks>
/// The client speaks IPP/2.0 and falls back to IPP/1.1 when the printer reports that 2.0 is unsupported.
/// It does not follow HTTP redirects and does not authenticate. Instances are safe for concurrent use.
/// </remarks>
public sealed class IppPrintClient : IDisposable {
    private const int StatusVersionNotSupported = 0x0503;
    private static readonly string[] JobAttributes = {
        "job-id", "job-uri", "job-state", "job-state-reasons", "job-state-message", "job-name"
    };

    private readonly IppHttpTransport _transport;
    private readonly IppPrintClientOptions _options;
    private readonly string _printerUri;
    private int _requestId;
    private volatile bool _useIpp11;

    /// <summary>Creates a client with its own HTTP connection handler.</summary>
    /// <param name="printerUri">Printer URI, such as ipp://printer.local/ipp/print. Port 631 is used when none is given.</param>
    /// <param name="options">Connection settings, or the defaults.</param>
    public IppPrintClient(Uri printerUri, IppPrintClientOptions? options = null)
        : this(printerUri, IppHttpTransport.CreateDefaultHandler(), true, options) {
    }

    /// <summary>Creates a client over a caller-supplied HTTP handler, for example to trust a printer's self-signed certificate.</summary>
    /// <param name="printerUri">Printer URI, such as ipps://printer.local/ipp/print.</param>
    /// <param name="handler">HTTP handler; it is not disposed with the client. Redirect responses are refused even if the handler would follow them.</param>
    /// <param name="options">Connection settings, or the defaults.</param>
    public IppPrintClient(Uri printerUri, HttpMessageHandler handler, IppPrintClientOptions? options = null)
        : this(printerUri, handler ?? throw new ArgumentNullException(nameof(handler)), false, options) {
    }

    private IppPrintClient(Uri printerUri, HttpMessageHandler handler, bool disposeHandler, IppPrintClientOptions? options) {
        ArgumentNullException.ThrowIfNull(printerUri);
        if (!printerUri.IsAbsoluteUri || (printerUri.Scheme != "ipp" && printerUri.Scheme != "ipps"))
            throw new ArgumentException("The printer URI must be an absolute ipp:// or ipps:// URI.", nameof(printerUri));
        if (string.IsNullOrEmpty(printerUri.Host) || !string.IsNullOrEmpty(printerUri.UserInfo) || !string.IsNullOrEmpty(printerUri.Fragment))
            throw new ArgumentException("The printer URI needs a host and cannot contain credentials or a fragment.", nameof(printerUri));
        _options = (options ?? new IppPrintClientOptions()).Snapshot();
        PrinterUri = printerUri;
        _printerUri = printerUri.AbsoluteUri;
        if (_printerUri.Length > 1023) throw new ArgumentException("The printer URI is too long.", nameof(printerUri));
        _transport = new IppHttpTransport(printerUri, handler, disposeHandler);
    }

    /// <summary>The printer URI sent in each request.</summary>
    public Uri PrinterUri { get; }

    /// <summary>Reads the printer's supported formats, media, finishing options and current state (Get-Printer-Attributes).</summary>
    public async Task<IppPrinterCapabilities> GetCapabilitiesAsync(CancellationToken cancellationToken = default) {
        IppMessage request = NewRequest(IppOperation.GetPrinterAttributes);
        request.Groups[0].Add("requested-attributes", Keywords("all", "media-col-database"));
        IppMessage response = await SendAsync(request, null, null, cancellationToken).ConfigureAwait(false);
        IppAttributeGroup printer = response.FirstGroup(IppGroupTag.Printer)
            ?? throw IppCodec.Malformed("the response has no printer attributes.");
        return IppAttributeMapper.Capabilities(printer);
    }

    /// <summary>Checks job attributes without printing (Validate-Job).</summary>
    /// <returns>Attributes the printer would ignore or substitute.</returns>
    /// <exception cref="IppException">The printer rejected the attributes or format.</exception>
    public async Task<IReadOnlyList<string>> ValidateJobAsync(string documentFormat, IppJobOptions? options = null, CancellationToken cancellationToken = default) {
        IppMessage request = JobRequest(IppOperation.ValidateJob, documentFormat, options);
        IppMessage response = await SendAsync(request, null, null, cancellationToken).ConfigureAwait(false);
        return IppAttributeMapper.UnsupportedNames(response);
    }

    /// <summary>Submits a document (Print-Job). Document bytes are streamed from the array.</summary>
    /// <inheritdoc cref="PrintAsync(Stream, string, IppJobOptions?, CancellationToken)"/>
    public Task<IppSubmission> PrintAsync(byte[] document, string documentFormat, IppJobOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        return PrintAsync(new MemoryStream(document, writable: false), documentFormat, options, cancellationToken);
    }

    /// <summary>Submits a document (Print-Job), streaming it from its current position.</summary>
    /// <param name="document">Document data. Seekable streams are sent with a length; others use chunked transfer. The stream is not disposed.</param>
    /// <param name="documentFormat">MIME type the printer lists in document-format-supported, such as application/pdf or image/pwg-raster.</param>
    /// <param name="options">Job attributes, or the printer defaults.</param>
    /// <param name="cancellationToken">Cancels the submission.</param>
    /// <returns>The accepted job. Acceptance is not completion.</returns>
    /// <exception cref="IppException">The printer was unreachable or rejected the job before accepting any of it.</exception>
    /// <exception cref="PdfPrintDeliveryException">The connection failed, timed out or was canceled after the document started streaming, or the acknowledgement was unusable; the job may print. Check the printer's jobs before retrying.</exception>
    public async Task<IppSubmission> PrintAsync(Stream document, string documentFormat, IppJobOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        if (!document.CanRead) throw new ArgumentException("The document stream must be readable.", nameof(document));
        IppMessage request = JobRequest(IppOperation.PrintJob, documentFormat, options);
        bool started = false;
        IppMessage response;
        try {
            response = await SendAsync(request, document, () => started = true, cancellationToken).ConfigureAwait(false);
        } catch (Exception error) when (started && (error is OperationCanceledException || error is IppException { IsDefinitiveRejection: false })) {
            throw new PdfPrintDeliveryException(null, error);
        }
        IppAttributeGroup? job = response.FirstGroup(IppGroupTag.Job);
        if (job?["job-id"]?.First is not { Tag: IppValueTag.Integer } id) {
            throw new PdfPrintDeliveryException(null, IppCodec.Malformed("the printer accepted the job without returning a job-id."));
        }
        return new IppSubmission(IppAttributeMapper.Job(job, id.Integer), IppAttributeMapper.UnsupportedNames(response));
    }

    /// <summary>Reads a job's state (Get-Job-Attributes).</summary>
    /// <exception cref="IppException">The printer does not know the job (status 0x0406) or the request failed.</exception>
    public async Task<IppJobInfo> GetJobAsync(int jobId, CancellationToken cancellationToken = default) {
        if (jobId < 1) throw new ArgumentOutOfRangeException(nameof(jobId));
        IppMessage request = NewRequest(IppOperation.GetJobAttributes);
        request.Groups[0].Add("job-id", IppValue.OfInteger(jobId)).Add("requested-attributes", Keywords(JobAttributes));
        IppMessage response = await SendAsync(request, null, null, cancellationToken).ConfigureAwait(false);
        IppAttributeGroup job = response.FirstGroup(IppGroupTag.Job) ?? throw IppCodec.Malformed("the response has no job attributes.");
        return IppAttributeMapper.Job(job, jobId);
    }

    /// <summary>Lists jobs the printer reports (Get-Jobs).</summary>
    /// <param name="which">Active or retained terminal jobs.</param>
    /// <param name="limit">Maximum number of jobs, from 1 through 1000.</param>
    /// <param name="cancellationToken">Cancels the request.</param>
    public async Task<IReadOnlyList<IppJobInfo>> GetJobsAsync(IppWhichJobs which = IppWhichJobs.NotCompleted, int limit = 100, CancellationToken cancellationToken = default) {
        if (!OfficeRuntimeCompatibility.IsDefined(which)) throw new ArgumentOutOfRangeException(nameof(which));
        if (limit is < 1 or > 1000) throw new ArgumentOutOfRangeException(nameof(limit));
        IppMessage request = NewRequest(IppOperation.GetJobs);
        request.Groups[0]
            .Add("limit", IppValue.OfInteger(limit))
            .Add("which-jobs", IppValue.OfString(IppValueTag.Keyword, which == IppWhichJobs.Completed ? "completed" : "not-completed"))
            .Add("requested-attributes", Keywords(JobAttributes));
        IppMessage response = await SendAsync(request, null, null, cancellationToken).ConfigureAwait(false);
        var jobs = new List<IppJobInfo>();
        foreach (IppAttributeGroup group in response.GroupsOf(IppGroupTag.Job)) {
            if (group["job-id"]?.First is { Tag: IppValueTag.Integer } id && id.Integer > 0) jobs.Add(IppAttributeMapper.Job(group, id.Integer));
        }
        return jobs;
    }

    /// <summary>Cancels a job (Cancel-Job).</summary>
    /// <exception cref="IppException">The job does not exist, is already terminal (status 0x0407), or the request failed.</exception>
    public async Task CancelJobAsync(int jobId, CancellationToken cancellationToken = default) {
        if (jobId < 1) throw new ArgumentOutOfRangeException(nameof(jobId));
        IppMessage request = NewRequest(IppOperation.CancelJob);
        request.Groups[0].Add("job-id", IppValue.OfInteger(jobId));
        await SendAsync(request, null, null, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Polls a job until it is canceled, aborted or completed.</summary>
    /// <param name="jobId">Job to follow.</param>
    /// <param name="timeout">Longest wait, or <see cref="Timeout.InfiniteTimeSpan"/>.</param>
    /// <param name="pollInterval">Delay between queries; defaults to one second.</param>
    /// <param name="cancellationToken">Stops waiting; the job is not canceled.</param>
    /// <returns>The job in its terminal state.</returns>
    /// <exception cref="TimeoutException">The job was not terminal when the timeout elapsed.</exception>
    public async Task<IppJobInfo> WaitForJobAsync(int jobId, TimeSpan timeout, TimeSpan? pollInterval = null, CancellationToken cancellationToken = default) {
        if (jobId < 1) throw new ArgumentOutOfRangeException(nameof(jobId));
        if (timeout < TimeSpan.Zero && timeout != Timeout.InfiniteTimeSpan) throw new ArgumentOutOfRangeException(nameof(timeout));
        TimeSpan interval = pollInterval ?? TimeSpan.FromSeconds(1);
        if (interval <= TimeSpan.Zero || interval > TimeSpan.FromHours(1)) throw new ArgumentOutOfRangeException(nameof(pollInterval));
        var clock = System.Diagnostics.Stopwatch.StartNew();
        while (true) {
            IppJobInfo job = await GetJobAsync(jobId, cancellationToken).ConfigureAwait(false);
            if (job.IsTerminal) return job;
            TimeSpan delay = interval;
            if (timeout != Timeout.InfiniteTimeSpan) {
                TimeSpan remaining = timeout - clock.Elapsed;
                if (remaining <= TimeSpan.Zero) {
                    throw new TimeoutException($"Job {jobId} was still {job.State} after {timeout}.");
                }
                if (remaining < delay) delay = remaining;
            }
            await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <inheritdoc />
    public void Dispose() => _transport.Dispose();

    private IppMessage JobRequest(IppOperation operation, string documentFormat, IppJobOptions? options) {
        if (documentFormat is null || documentFormat.Length is 0 or > 255 || documentFormat.IndexOf('/') <= 0 || documentFormat.Any(c => c <= ' ' || c > '~'))
            throw new ArgumentException("A MIME media type such as application/pdf is required.", nameof(documentFormat));
        IppJobOptions job = (options ?? new IppJobOptions()).Snapshot();
        IppMessage request = NewRequest(operation);
        request.Groups[0]
            .Add("job-name", IppValue.OfString(IppValueTag.NameWithoutLanguage, job.JobName))
            .Add("ipp-attribute-fidelity", IppValue.OfBoolean(job.AttributeFidelity))
            .Add("document-format", IppValue.OfString(IppValueTag.MimeMediaType, documentFormat));
        IppAttributeMapper.AddJobTemplate(request, job);
        return request;
    }

    private IppMessage NewRequest(IppOperation operation) {
        var request = new IppMessage(2, 0, (short)operation, 0);
        IppAttributeGroup group = request.AddGroup(IppGroupTag.Operation)
            .Add("attributes-charset", IppValue.OfString(IppValueTag.Charset, "utf-8"))
            .Add("attributes-natural-language", IppValue.OfString(IppValueTag.NaturalLanguage, "en"))
            .Add("printer-uri", IppValue.OfString(IppValueTag.Uri, _printerUri));
        if (_options.RequestingUserName is { } user) group.Add("requesting-user-name", IppValue.OfString(IppValueTag.NameWithoutLanguage, user));
        return request;
    }

    private async Task<IppMessage> SendAsync(IppMessage request, Stream? document, Action? bodyStarted, CancellationToken cancellationToken) {
        long documentStart = document is { CanSeek: true } ? document.Position : -1;
        while (true) {
            if (_useIpp11) {
                request.VersionMajor = 1;
                request.VersionMinor = 1;
            }
            request.RequestId = (Interlocked.Increment(ref _requestId) & int.MaxValue) | 1;
            IppMessage response = await _transport.SendAsync(request, document, bodyStarted, _options.Timeout, _options.MaxResponseBytes, cancellationToken).ConfigureAwait(false);
            int status = (ushort)response.Code;
            if (status < 0x0100) return response;
            // A version error rejects the request before any processing, so it can be replayed once with IPP/1.1
            // when the document (if any) can be rewound.
            if (status == StatusVersionNotSupported && !_useIpp11 && (document is null || documentStart >= 0)) {
                _useIpp11 = true;
                if (document is not null) document.Position = documentStart;
                continue;
            }
            string? message = response.FirstGroup(IppGroupTag.Operation)?["status-message"]?.First?.Text;
            throw new IppException(
                $"The printer rejected {(IppOperation)request.Code} with status 0x{status:X4} ({IppAttributeMapper.StatusName(status)})"
                + (string.IsNullOrEmpty(message) ? "." : ": " + message), status);
        }
    }

    private static IppValue[] Keywords(params string[] values) =>
        values.Select(value => IppValue.OfString(IppValueTag.Keyword, value)).ToArray();
}

/// <summary>Translates between IPP attributes and the public printing types.</summary>
internal static class IppAttributeMapper {
    public static IppPrinterCapabilities Capabilities(IppAttributeGroup printer) {
        int? maxCopies = printer["copies-supported"]?.Values.Where(v => v.Tag == IppValueTag.RangeOfInteger).Select(v => (int?)v.Second).FirstOrDefault();
        return new IppPrinterCapabilities {
            MakeAndModel = Text(printer["printer-make-and-model"]),
            State = EnumValue<IppPrinterState>(printer["printer-state"]) ?? IppPrinterState.Idle,
            StateReasons = Strings(printer["printer-state-reasons"]).Where(reason => reason != "none").ToArray(),
            StateMessage = Text(printer["printer-state-message"]),
            DocumentFormats = Strings(printer["document-format-supported"]),
            Media = Strings(printer["media-supported"]),
            MediaDefault = Text(printer["media-default"]),
            MediaSizes = (printer["media-col-database"]?.Values ?? new List<IppValue>()).Select(Media).OfType<IppMedia>().ToArray(),
            MediaColDefault = printer["media-col-default"]?.First is { } mediaDefault ? Media(mediaDefault) : null,
            Sides = Strings(printer["sides-supported"]),
            SidesDefault = Text(printer["sides-default"]),
            ColorModes = Strings(printer["print-color-mode-supported"]),
            ColorModeDefault = Text(printer["print-color-mode-default"]),
            Resolutions = Resolutions(printer["printer-resolution-supported"]),
            ResolutionDefault = Resolutions(printer["printer-resolution-default"]).FirstOrDefault(),
            MaxCopies = maxCopies is > 0 ? maxCopies.Value : 1,
            Orientations = (printer["orientation-requested-supported"]?.Values ?? new List<IppValue>())
                .Where(v => v.Tag == IppValueTag.Enum && v.Integer is >= 3 and <= 6).Select(v => (IppOrientation)v.Integer).ToArray(),
            PrintScaling = Strings(printer["print-scaling-supported"]),
            PageRangesSupported = printer["page-ranges-supported"]?.First is { Tag: IppValueTag.Boolean, Boolean: true },
            PwgRasterResolutions = Resolutions(printer["pwg-raster-document-resolution-supported"]),
            PwgRasterTypes = Strings(printer["pwg-raster-document-type-supported"]),
            IppVersions = Strings(printer["ipp-versions-supported"])
        };
    }

    public static IppJobInfo Job(IppAttributeGroup job, int id) {
        int? state = job["job-state"]?.First is { Tag: IppValueTag.Enum } value ? value.Integer : null;
        return new IppJobInfo(id, state is >= 3 and <= 9 ? (IppJobState)state.Value : IppJobState.Pending) {
            Uri = Text(job["job-uri"]),
            StateReasons = Strings(job["job-state-reasons"]).Where(reason => reason != "none").ToArray(),
            StateMessage = Text(job["job-state-message"]),
            Name = Text(job["job-name"])
        };
    }

    public static IReadOnlyList<string> UnsupportedNames(IppMessage response) =>
        response.GroupsOf(IppGroupTag.Unsupported).SelectMany(group => group.Attributes).Select(attribute => attribute.Name).Distinct(StringComparer.Ordinal).ToArray();

    public static void AddJobTemplate(IppMessage request, IppJobOptions job) {
        var group = new IppAttributeGroup(IppGroupTag.Job);
        if (job.Copies is { } copies) group.Add("copies", IppValue.OfInteger(copies));
        if (job.Sides is { } sides) group.Add("sides", IppValue.OfString(IppValueTag.Keyword, sides));
        if (job.Media is { } media) group.Add("media", IppValue.OfString(IppValueTag.Keyword, media));
        if (job.MediaCol is { } mediaCol) group.Add("media-col", MediaCol(mediaCol));
        if (job.ColorMode is { } colorMode) group.Add("print-color-mode", IppValue.OfString(IppValueTag.Keyword, colorMode));
        if (job.Orientation is { } orientation) group.Add("orientation-requested", IppValue.OfEnum((int)orientation));
        if (job.PrintScaling is { } scaling) group.Add("print-scaling", IppValue.OfString(IppValueTag.Keyword, scaling));
        if (job.PageRanges is { } ranges) group.Add("page-ranges", ranges.Select(range => IppValue.OfRange(range.First, range.Last)).ToArray());
        if (job.Resolution is { } resolution) group.Add("printer-resolution", IppValue.OfResolution(resolution.CrossFeed, resolution.Feed, (byte)resolution.Units));
        if (group.Attributes.Count > 0) request.Groups.Add(group);
    }

    public static IppValue MediaCol(IppMedia media) {
        var members = new List<IppAttribute> {
            new("media-size", IppValue.OfCollection(new[] {
                new IppAttribute("x-dimension", IppValue.OfInteger(media.Width)),
                new IppAttribute("y-dimension", IppValue.OfInteger(media.Height))
            }))
        };
        if (media.SizeName is { } name) members.Add(new("media-size-name", IppValue.OfString(IppValueTag.Keyword, name)));
        if (media.Source is { } source) members.Add(new("media-source", IppValue.OfString(IppValueTag.Keyword, source)));
        if (media.Type is { } type) members.Add(new("media-type", IppValue.OfString(IppValueTag.Keyword, type)));
        if (media.LeftMargin is { } left) members.Add(new("media-left-margin", IppValue.OfInteger(left)));
        if (media.TopMargin is { } top) members.Add(new("media-top-margin", IppValue.OfInteger(top)));
        if (media.RightMargin is { } right) members.Add(new("media-right-margin", IppValue.OfInteger(right)));
        if (media.BottomMargin is { } bottom) members.Add(new("media-bottom-margin", IppValue.OfInteger(bottom)));
        return IppValue.OfCollection(members);
    }

    /// <summary>Reads a media-col value with a fixed media-size; null for custom ranges or incomplete entries.</summary>
    public static IppMedia? Media(IppValue value) {
        if (value.Members is not { } members) return null;
        IppValue? size = Member(members, "media-size");
        if (size?.Members is not { } dimensions) return null;
        if (Member(dimensions, "x-dimension") is not { Tag: IppValueTag.Integer, Integer: > 0 } x
            || Member(dimensions, "y-dimension") is not { Tag: IppValueTag.Integer, Integer: > 0 } y) return null;
        return new IppMedia(x.Integer, y.Integer) {
            SizeName = Member(members, "media-size-name")?.Text ?? Member(members, "media-key")?.Text,
            Source = Member(members, "media-source")?.Text,
            Type = Member(members, "media-type")?.Text,
            LeftMargin = Integer(Member(members, "media-left-margin")),
            TopMargin = Integer(Member(members, "media-top-margin")),
            RightMargin = Integer(Member(members, "media-right-margin")),
            BottomMargin = Integer(Member(members, "media-bottom-margin"))
        };
    }

    public static string StatusName(int status) => status switch {
        0x0400 => "client-error-bad-request",
        0x0401 => "client-error-forbidden",
        0x0402 => "client-error-not-authenticated",
        0x0403 => "client-error-not-authorized",
        0x0404 => "client-error-not-possible",
        0x0405 => "client-error-timeout",
        0x0406 => "client-error-not-found",
        0x0407 => "client-error-gone",
        0x0408 => "client-error-request-entity-too-large",
        0x040A => "client-error-document-format-not-supported",
        0x040B => "client-error-attributes-or-values-not-supported",
        0x040D => "client-error-conflicting-attributes",
        0x040E => "client-error-compression-not-supported",
        0x0410 => "client-error-document-format-error",
        0x0500 => "server-error-internal-error",
        0x0501 => "server-error-operation-not-supported",
        0x0502 => "server-error-service-unavailable",
        0x0503 => "server-error-version-not-supported",
        0x0506 => "server-error-not-accepting-jobs",
        0x0507 => "server-error-busy",
        0x0508 => "server-error-job-canceled",
        _ => status >= 0x0500 ? "server error" : "client error"
    };

    private static IppValue? Member(IReadOnlyList<IppAttribute> members, string name) {
        foreach (IppAttribute member in members) {
            if (member.Name == name) return member.First;
        }
        return null;
    }

    private static int? Integer(IppValue? value) => value is { Tag: IppValueTag.Integer } ? value.Integer : null;

    private static string? Text(IppAttribute? attribute) => attribute?.Values.FirstOrDefault(value => value.Text is not null)?.Text;

    private static IReadOnlyList<string> Strings(IppAttribute? attribute) =>
        attribute is null ? Array.Empty<string>() : attribute.Values.Where(value => value.Text is not null).Select(value => value.Text!).ToArray();

    private static T? EnumValue<T>(IppAttribute? attribute) where T : struct, System.Enum {
        if (attribute?.First is not { Tag: IppValueTag.Enum } value) return null;
        T candidate = (T)(object)value.Integer;
        return OfficeRuntimeCompatibility.IsDefined(candidate) ? candidate : null;
    }

    private static IReadOnlyList<IppResolution> Resolutions(IppAttribute? attribute) =>
        attribute is null ? Array.Empty<IppResolution>() : attribute.Values
            .Where(value => value.Tag == IppValueTag.Resolution && value.Integer > 0 && value.Second > 0 && value.Units is 3 or 4)
            .Select(value => new IppResolution(value.Integer, value.Second, (IppResolutionUnits)value.Units)).ToArray();
}
