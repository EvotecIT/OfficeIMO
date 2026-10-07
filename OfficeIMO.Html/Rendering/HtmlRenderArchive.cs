using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Limits for deterministic PNG or SVG page archive packaging.</summary>
public sealed class HtmlRenderArchiveOptions {
    /// <summary>Default maximum final ZIP size.</summary>
    public const long DefaultMaximumArchiveBytes = 1024L * 1024L * 1024L;
    /// <summary>Default maximum UTF-8 manifest size.</summary>
    public const int DefaultMaximumManifestBytes = 4 * 1024 * 1024;

    /// <summary>Maximum final ZIP size, including entries, manifest, and ZIP metadata.</summary>
    public long MaximumArchiveBytes { get; set; } = DefaultMaximumArchiveBytes;
    /// <summary>Maximum UTF-8 manifest size.</summary>
    public int MaximumManifestBytes { get; set; } = DefaultMaximumManifestBytes;

    internal HtmlRenderArchiveOptions Clone() => new HtmlRenderArchiveOptions {
        MaximumArchiveBytes = MaximumArchiveBytes,
        MaximumManifestBytes = MaximumManifestBytes
    };

    internal void Validate() {
        if (MaximumArchiveBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumArchiveBytes));
        if (MaximumManifestBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumManifestBytes));
    }
}

/// <summary>Thrown when page archive packaging exceeds a caller-controlled byte limit.</summary>
public sealed class HtmlRenderArchiveLimitException : InvalidOperationException {
    internal HtmlRenderArchiveLimitException(string limitName, long attemptedBytes, long maximumBytes)
        : base("HTML render archive " + limitName + " requires at least " +
               attemptedBytes.ToString(System.Globalization.CultureInfo.InvariantCulture) +
               " bytes, exceeding the configured limit of " +
               maximumBytes.ToString(System.Globalization.CultureInfo.InvariantCulture) + " bytes.") {
        LimitName = limitName;
        AttemptedBytes = attemptedBytes;
        MaximumBytes = maximumBytes;
    }

    /// <summary>Name of the exceeded limit.</summary>
    public string LimitName { get; }
    /// <summary>First observed size beyond the limit.</summary>
    public long AttemptedBytes { get; }
    /// <summary>Configured maximum size.</summary>
    public long MaximumBytes { get; }
}

/// <summary>One separately encoded page described by an HTML render archive manifest.</summary>
public sealed class HtmlRenderArchivePage {
    private readonly ReadOnlyCollection<HtmlRenderSourcePlacement> _sourcePlacements;
    private readonly ReadOnlyCollection<OfficeImageExportDiagnostic> _encodingDiagnostics;

    internal HtmlRenderArchivePage(HtmlRenderSurface surface, OfficeImageExportResult image,
        string entryName, string sha256, IEnumerable<OfficeImageExportDiagnostic> encodingDiagnostics) {
        OutputIndex = surface.OutputIndex;
        SourcePageNumber = surface.Descriptor.SourcePageNumber;
        Width = surface.Descriptor.Width;
        Height = surface.Descriptor.Height;
        IsClipped = surface.Descriptor.IsClipped;
        EncodedWidth = image.Width;
        EncodedHeight = image.Height;
        MimeType = image.MimeType;
        EncodedLength = image.EncodedLength;
        EntryName = entryName;
        Sha256 = sha256;
        _sourcePlacements = new List<HtmlRenderSourcePlacement>(surface.Descriptor.SourcePlacements).AsReadOnly();
        _encodingDiagnostics = new List<OfficeImageExportDiagnostic>(
            encodingDiagnostics ?? throw new ArgumentNullException(nameof(encodingDiagnostics))).AsReadOnly();
    }

    /// <summary>Zero-based output order.</summary>
    public int OutputIndex { get; }
    /// <summary>One-based primary source page number.</summary>
    public int SourcePageNumber { get; }
    /// <summary>Retained surface width in CSS pixels.</summary>
    public double Width { get; }
    /// <summary>Retained surface height in CSS pixels.</summary>
    public double Height { get; }
    /// <summary>Whether the retained surface clips or composes source content.</summary>
    public bool IsClipped { get; }
    /// <summary>Encoded width in pixels for PNG or CSS pixels for SVG.</summary>
    public int EncodedWidth { get; }
    /// <summary>Encoded height in pixels for PNG or CSS pixels for SVG.</summary>
    public int EncodedHeight { get; }
    /// <summary>Canonical media type.</summary>
    public string MimeType { get; }
    /// <summary>Encoded payload length.</summary>
    public long EncodedLength { get; }
    /// <summary>Safe relative ZIP entry name.</summary>
    public string EntryName { get; }
    /// <summary>Lowercase SHA-256 digest of the exact encoded entry.</summary>
    public string Sha256 { get; }
    /// <summary>Whether encoding this page reported an approximation, omission, or failure.</summary>
    public bool HasLoss => _encodingDiagnostics.Any(diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <summary>Scale, font, codec, and other diagnostics produced while encoding this page.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> EncodingDiagnostics => _encodingDiagnostics;
    /// <summary>Every source page or slice contributing to this encoded page.</summary>
    public IReadOnlyList<HtmlRenderSourcePlacement> SourcePlacements => _sourcePlacements;
}

/// <summary>Deterministic manifest retained with a separately encoded HTML page archive.</summary>
public sealed class HtmlRenderArchiveManifest {
    /// <summary>Stable manifest schema identifier.</summary>
    public const string SchemaId = "officeimo.html.render-archive";
    /// <summary>Current manifest schema version.</summary>
    public const int SchemaVersion = 1;

    private readonly ReadOnlyCollection<HtmlRenderArchivePage> _pages;
    private readonly ReadOnlyCollection<string> _providerIds;
    private readonly ReadOnlyCollection<HtmlDiagnostic> _diagnostics;

    internal HtmlRenderArchiveManifest(HtmlRenderResult result, OfficeImageExportFormat format,
        IEnumerable<HtmlRenderArchivePage> pages) {
        ProfileId = result.Request.ProfileId;
        DocumentState = result.Request.DocumentState;
        CssMedia = result.Request.CssMedia;
        Surface = result.Request.Surface;
        Pagination = result.Request.Pagination;
        PageSet = result.Request.PageSet.Mode;
        FirstPageIndex = result.Request.PageSet.FirstPageIndex;
        PageCount = result.Request.PageSet.PageCount;
        Encoder = result.Request.Encoder;
        ImageFormat = format;
        Coverage = result.Coverage;
        RequestedScale = result.RequestedScale;
        BackgroundColor = result.BackgroundColor;
        _pages = new List<HtmlRenderArchivePage>(pages).AsReadOnly();
        HasLoss = result.HasLoss || _pages.Any(page => page.HasLoss);
        _providerIds = new List<string>(result.DeclaredProviderIds).AsReadOnly();
        _diagnostics = new List<HtmlDiagnostic>(result.Diagnostics).AsReadOnly();
    }

    /// <summary>Versioned render profile identifier.</summary>
    public string ProfileId { get; }
    /// <summary>Document snapshot provenance.</summary>
    public HtmlRenderDocumentState DocumentState { get; }
    /// <summary>CSS media axis.</summary>
    public HtmlCssMediaContext CssMedia { get; }
    /// <summary>Layout surface axis.</summary>
    public HtmlRenderLayoutSurface Surface { get; }
    /// <summary>Pagination axis.</summary>
    public HtmlRenderPaginationPolicy Pagination { get; }
    /// <summary>Page selection and composition axis.</summary>
    public HtmlRenderPageSetMode PageSet { get; }
    /// <summary>Zero-based first source page requested by the selection.</summary>
    public int FirstPageIndex { get; }
    /// <summary>Requested range length, or null when every remaining page was requested.</summary>
    public int? PageCount { get; }
    /// <summary>Requested render encoder.</summary>
    public HtmlRenderEncoder Encoder { get; }
    /// <summary>Resolved shared image format.</summary>
    public OfficeImageExportFormat ImageFormat { get; }
    /// <summary>Qualification of the exact effective request axes.</summary>
    public HtmlCapabilityCoverage Coverage { get; }
    /// <summary>Requested scale before bounded raster reduction.</summary>
    public double RequestedScale { get; }
    /// <summary>Requested image background.</summary>
    public OfficeColor BackgroundColor { get; }
    /// <summary>Whether diagnostics report an approximation, omission, or failure.</summary>
    public bool HasLoss { get; }
    /// <summary>Separately encoded pages in output order.</summary>
    public IReadOnlyList<HtmlRenderArchivePage> Pages => _pages;
    /// <summary>Provider identities declared by the profile compatibility manifests.</summary>
    public IReadOnlyList<string> ProviderIds => _providerIds;
    /// <summary>Structured parse, layout, resource, and fallback diagnostics.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Serializes this manifest using stable property and collection order.</summary>
    public string ToJson() => HtmlRenderArchiveManifestJsonWriter.ToJson(this);
}

/// <summary>Completed deterministic HTML page archive and its typed manifest.</summary>
public sealed class HtmlRenderArchiveResult {
    private readonly byte[] _bytes;

    internal HtmlRenderArchiveResult(byte[] bytes, HtmlRenderArchiveManifest manifest) {
        _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
        Manifest = manifest ?? throw new ArgumentNullException(nameof(manifest));
    }

    /// <summary>Typed manifest embedded as <c>manifest.json</c>.</summary>
    public HtmlRenderArchiveManifest Manifest { get; }
    /// <summary>Exact ZIP byte count without copying the payload.</summary>
    public long EncodedLength => _bytes.LongLength;
    /// <summary>Detached copy of the ZIP bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Writes the ZIP payload to a caller-owned stream without closing it.</summary>
    public HtmlRenderArchiveResult Save(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        stream.Write(_bytes, 0, _bytes.Length);
        return this;
    }

    /// <summary>Writes the ZIP payload asynchronously without closing the caller-owned stream.</summary>
    public async Task<HtmlRenderArchiveResult> SaveAsync(Stream stream,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        await stream.WriteAsync(_bytes, 0, _bytes.Length, cancellationToken).ConfigureAwait(false);
        return this;
    }
}
