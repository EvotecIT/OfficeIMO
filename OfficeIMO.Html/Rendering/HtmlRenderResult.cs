using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Placement of one source surface or source slice in a retained output surface.</summary>
public sealed class HtmlRenderSourcePlacement {
    internal HtmlRenderSourcePlacement(int sourcePageNumber, double sourceOffsetX, double sourceOffsetY,
        double width, double height, double outputOffsetX, double outputOffsetY, bool clipped) {
        SourcePageNumber = sourcePageNumber;
        SourceOffsetX = sourceOffsetX;
        SourceOffsetY = sourceOffsetY;
        Width = width;
        Height = height;
        OutputOffsetX = outputOffsetX;
        OutputOffsetY = outputOffsetY;
        IsClipped = clipped;
    }

    /// <summary>One-based source page number.</summary>
    public int SourcePageNumber { get; }
    /// <summary>Horizontal origin within the source page or continuous canvas.</summary>
    public double SourceOffsetX { get; }
    /// <summary>Vertical origin within the source page or continuous canvas.</summary>
    public double SourceOffsetY { get; }
    /// <summary>Placed width in CSS pixels.</summary>
    public double Width { get; }
    /// <summary>Placed height in CSS pixels.</summary>
    public double Height { get; }
    /// <summary>Horizontal origin in the output surface.</summary>
    public double OutputOffsetX { get; }
    /// <summary>Vertical origin in the output surface.</summary>
    public double OutputOffsetY { get; }
    /// <summary>Whether the source contribution was clipped.</summary>
    public bool IsClipped { get; }
    /// <summary>Source rectangle contributing to the output surface.</summary>
    public HtmlRenderRectangle SourceBounds => new HtmlRenderRectangle(SourceOffsetX, SourceOffsetY, Width, Height);
    /// <summary>Rectangle occupied by the contribution in the output surface.</summary>
    public HtmlRenderRectangle OutputBounds => new HtmlRenderRectangle(OutputOffsetX, OutputOffsetY, Width, Height);
}

/// <summary>One output surface retained by a resolved HTML render request.</summary>
public sealed class HtmlRenderSurfaceResult {
    private readonly ReadOnlyCollection<HtmlRenderSourcePlacement> _sourcePlacements;

    internal HtmlRenderSurfaceResult(int outputIndex, int sourcePageNumber, double width, double height,
        double sourceOffsetX, double sourceOffsetY, bool clipped,
        IEnumerable<HtmlRenderSourcePlacement>? sourcePlacements = null) {
        OutputIndex = outputIndex;
        SourcePageNumber = sourcePageNumber;
        Width = width;
        Height = height;
        SourceOffsetX = sourceOffsetX;
        SourceOffsetY = sourceOffsetY;
        IsClipped = clipped;
        _sourcePlacements = new List<HtmlRenderSourcePlacement>(sourcePlacements ?? new[] {
            new HtmlRenderSourcePlacement(sourcePageNumber, sourceOffsetX, sourceOffsetY,
                width, height, 0D, 0D, clipped)
        }).AsReadOnly();
        if (_sourcePlacements.Count == 0) throw new ArgumentException("At least one source placement is required.", nameof(sourcePlacements));
    }

    /// <summary>Zero-based position in the resolved output set.</summary>
    public int OutputIndex { get; }
    /// <summary>One-based page number in the preselection render.</summary>
    public int SourcePageNumber { get; }
    /// <summary>Surface width in CSS pixels.</summary>
    public double Width { get; }
    /// <summary>Surface height in CSS pixels.</summary>
    public double Height { get; }
    /// <summary>Horizontal offset into the original completed layout.</summary>
    public double SourceOffsetX { get; }
    /// <summary>Vertical offset into the original completed layout.</summary>
    public double SourceOffsetY { get; }
    /// <summary>Whether this surface clips a viewport, page slice, or stitched source page.</summary>
    public bool IsClipped { get; }
    /// <summary>Every source page or slice contributing to this output surface.</summary>
    public IReadOnlyList<HtmlRenderSourcePlacement> SourcePlacements => _sourcePlacements;
    /// <summary>Complete output-surface rectangle.</summary>
    public HtmlRenderRectangle Bounds => new HtmlRenderRectangle(0D, 0D, Width, Height);

    /// <summary>Maps an output point to the contributing source page or canvas slice.</summary>
    public bool TryMapToSource(HtmlRenderPoint outputPoint, out HtmlRenderSourcePoint? sourcePoint) {
        for (int index = _sourcePlacements.Count - 1; index >= 0; index--) {
            HtmlRenderSourcePlacement placement = _sourcePlacements[index];
            if (!placement.OutputBounds.Contains(outputPoint)) continue;
            sourcePoint = new HtmlRenderSourcePoint(
                OutputIndex,
                placement.SourcePageNumber,
                outputPoint,
                new HtmlRenderPoint(
                    placement.SourceOffsetX + outputPoint.X - placement.OutputOffsetX,
                    placement.SourceOffsetY + outputPoint.Y - placement.OutputOffsetY),
                placement.IsClipped);
            return true;
        }
        sourcePoint = null;
        return false;
    }

    /// <summary>Maps finite output coordinates to the contributing source page or canvas slice.</summary>
    public bool TryMapToSource(double outputX, double outputY, out HtmlRenderSourcePoint? sourcePoint) =>
        TryMapToSource(new HtmlRenderPoint(outputX, outputY), out sourcePoint);
}

/// <summary>
/// A resolved request and its retained backend-neutral scene. Output adapters consume this
/// result without choosing new CSS, layout, pagination, or page-selection defaults.
/// </summary>
public sealed class HtmlRenderResult {
    private readonly ReadOnlyCollection<HtmlRenderSurfaceResult> _surfaces;
    private readonly ReadOnlyCollection<HtmlRenderSurface> _outputSurfaces;
    private readonly ReadOnlyCollection<string> _declaredProviderIds;
    private readonly ReadOnlyCollection<HtmlDiagnostic> _lossDiagnostics;

    internal HtmlRenderResult(HtmlRenderRequest request, HtmlRenderDocument document,
        IEnumerable<HtmlRenderSurfaceResult> surfaces) {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        Document = document ?? throw new ArgumentNullException(nameof(document));
        _surfaces = new List<HtmlRenderSurfaceResult>(surfaces ?? throw new ArgumentNullException(nameof(surfaces))).AsReadOnly();
        if (_surfaces.Count != document.Pages.Count) {
            throw new ArgumentException("Every retained page requires one surface descriptor.", nameof(surfaces));
        }
        _outputSurfaces = document.Pages
            .Select((page, index) => new HtmlRenderSurface(this, page, _surfaces[index]))
            .ToList().AsReadOnly();
        HtmlRenderProfileContract profile = HtmlRenderProfileContracts.Get(request.Profile);
        _declaredProviderIds = (request.MatchesNamedProfile ? profile.CapabilityProfileIds : Array.Empty<string>())
            .SelectMany(profileId => HtmlRenderCapabilityCatalog.GetProfile(profileId).Providers)
            .Select(provider => provider.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(providerId => providerId, StringComparer.Ordinal)
            .ToList().AsReadOnly();
        _lossDiagnostics = document.Diagnostics
            .Where(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None)
            .ToList().AsReadOnly();
    }

    /// <summary>Independent immutable request snapshot used for this operation.</summary>
    public HtmlRenderRequest Request { get; }
    /// <summary>Selected, sliced, or stitched backend-neutral display list.</summary>
    public HtmlRenderDocument Document { get; }
    /// <summary>Ordered surface geometry and source mapping.</summary>
    public IReadOnlyList<HtmlRenderSurfaceResult> Surfaces => _surfaces;
    /// <summary>Ordered executable surface views for preview, drawing, geometry, and hit-test consumers.</summary>
    public IReadOnlyList<HtmlRenderSurface> OutputSurfaces => _outputSurfaces;
    /// <summary>Requested output scale before any bounded raster scale reduction.</summary>
    public double RequestedScale => Request.Options.Scale;
    /// <summary>Requested output background used by image adapters.</summary>
    public OfficeColor BackgroundColor => Request.Options.BackgroundColor;
    /// <summary>Qualification of the exact effective request axes.</summary>
    public HtmlCapabilityCoverage Coverage => Request.Coverage;
    /// <summary>Provider identities declared by the compatibility manifests for this profile.</summary>
    public IReadOnlyList<string> DeclaredProviderIds => _declaredProviderIds;
    /// <summary>Structured diagnostics emitted by parsing, layout, and display-list construction.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => Document.Diagnostics;
    /// <summary>Whether the retained result contains an approximation, omission, or failure.</summary>
    public bool HasLoss => Document.HasLoss;
    /// <summary>Diagnostics that report an approximation, omission, or failure.</summary>
    public IReadOnlyList<HtmlDiagnostic> LossDiagnostics => _lossDiagnostics;

    /// <summary>Gets one executable output surface by zero-based index.</summary>
    public HtmlRenderSurface GetSurface(int outputIndex) {
        if (outputIndex < 0 || outputIndex >= _outputSurfaces.Count) {
            throw new ArgumentOutOfRangeException(nameof(outputIndex));
        }
        return _outputSurfaces[outputIndex];
    }

    /// <summary>
    /// Returns an equivalent retained result carrying additional diagnostics from an owning
    /// container, adapter, or input boundary. The current result remains unchanged.
    /// </summary>
    public HtmlRenderResult WithAdditionalDiagnostics(IEnumerable<HtmlDiagnostic> diagnostics) {
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        return new HtmlRenderResult(Request, Document.WithAdditionalDiagnostics(diagnostics), _surfaces);
    }

    /// <summary>Throws with the complete report when the retained result is not lossless.</summary>
    public HtmlRenderResult RequireNoLoss() {
        Document.RequireNoLoss();
        return this;
    }
}
