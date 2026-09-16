namespace OfficeIMO.Html;

/// <summary>
/// Describes the deterministic media environment used to select one responsive image candidate.
/// </summary>
public sealed class HtmlResponsiveImageSelectionOptions {
    /// <summary>Viewport width in CSS pixels.</summary>
    public double ViewportWidth { get; set; } = 816D;

    /// <summary>Viewport height in CSS pixels.</summary>
    public double ViewportHeight { get; set; } = 1056D;

    /// <summary>Target device-pixel density. The default is one device pixel per CSS pixel.</summary>
    public double DevicePixelRatio { get; set; } = 1D;

    /// <summary>Media type used by conditions in the <c>sizes</c> list.</summary>
    public HtmlCssMediaContext MediaContext { get; set; } = HtmlCssMediaContext.Screen;

    /// <summary>Deterministic device and user-preference values used by media conditions.</summary>
    public HtmlRenderMediaFeatures MediaFeatures { get; set; } = new HtmlRenderMediaFeatures();

    /// <summary>Root and initial font size used to resolve relative lengths in <c>sizes</c>.</summary>
    public double DefaultFontSize { get; set; } = 16D;

    /// <summary>Optional cap applied while parsing source-set candidates.</summary>
    public int? MaxCandidates { get; set; } = HtmlConversionLimits.DefaultMaxResponsiveImageCandidates;

    /// <summary>Optional character cap applied before parsing a <c>sizes</c> value.</summary>
    public int? MaxSizesCharacters { get; set; } = HtmlConversionLimits.DefaultMaxResponsiveImageSizesCharacters;

    internal void Validate() {
        if (!HtmlResponsiveImageSelector.IsFinite(ViewportWidth) || ViewportWidth <= 0D) throw new ArgumentOutOfRangeException(nameof(ViewportWidth));
        if (!HtmlResponsiveImageSelector.IsFinite(ViewportHeight) || ViewportHeight <= 0D) throw new ArgumentOutOfRangeException(nameof(ViewportHeight));
        if (!HtmlResponsiveImageSelector.IsFinite(DevicePixelRatio) || DevicePixelRatio <= 0D) throw new ArgumentOutOfRangeException(nameof(DevicePixelRatio));
        if (!HtmlResponsiveImageSelector.IsFinite(DefaultFontSize) || DefaultFontSize <= 0D) throw new ArgumentOutOfRangeException(nameof(DefaultFontSize));
        if (MaxCandidates.HasValue && MaxCandidates.Value <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCandidates));
        if (MaxSizesCharacters.HasValue && MaxSizesCharacters.Value <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSizesCharacters));
        if (MediaContext is not (HtmlCssMediaContext.Screen or HtmlCssMediaContext.Print)) throw new ArgumentOutOfRangeException(nameof(MediaContext));
        (MediaFeatures ?? throw new ArgumentNullException(nameof(MediaFeatures))).Validate();
    }
}

/// <summary>Result of deterministic responsive image selection.</summary>
public readonly struct HtmlResponsiveImageSelection {
    internal HtmlResponsiveImageSelection(HtmlSrcSetCandidate candidate, double effectiveDensity, double sourceSize,
        bool usesDefaultSource) {
        Candidate = candidate;
        EffectiveDensity = effectiveDensity;
        SourceSize = sourceSize;
        UsesDefaultSource = usesDefaultSource;
        HasValue = !string.IsNullOrWhiteSpace(candidate.Url);
    }

    /// <summary>Whether a valid candidate was selected.</summary>
    public bool HasValue { get; }

    /// <summary>Selected candidate, preserving its authored descriptor.</summary>
    public HtmlSrcSetCandidate Candidate { get; }

    /// <summary>Candidate density after normalizing a width descriptor against the selected source size.</summary>
    public double EffectiveDensity { get; }

    /// <summary>Selected source size in CSS pixels. Density-only sets use the viewport width.</summary>
    public double SourceSize { get; }

    /// <summary>Whether the selected candidate was synthesized from the separate default source.</summary>
    public bool UsesDefaultSource { get; }
}

/// <summary>
/// Selects one candidate from an HTML <c>srcset</c> and <c>sizes</c> pair without depending on a DOM provider.
/// </summary>
public static class HtmlResponsiveImageSelector {
    /// <summary>
    /// Selects one source using deterministic density preference. Width descriptors are normalized by the
    /// first matching supported <c>sizes</c> entry; an omitted or unsupported list falls back to <c>100vw</c>.
    /// </summary>
    /// <param name="srcSet">Authored source-set value.</param>
    /// <param name="sizes">Optional source-size list.</param>
    /// <param name="defaultSource">Optional <c>src</c> fallback. It participates as <c>1x</c> only for density sets without an authored <c>1x</c> candidate.</param>
    /// <param name="options">Selection environment.</param>
    public static HtmlResponsiveImageSelection Select(string? srcSet, string? sizes, string? defaultSource,
        HtmlResponsiveImageSelectionOptions? options = null) {
        HtmlResponsiveImageSelectionOptions resolved = options ?? new HtmlResponsiveImageSelectionOptions();
        resolved.Validate();
        IReadOnlyList<HtmlSrcSetCandidate> parsed = HtmlSrcSetParser.Parse(srcSet, resolved.MaxCandidates);
        return Select(parsed, sizes, defaultSource, resolved);
    }

    internal static HtmlResponsiveImageSelection Select(IReadOnlyList<HtmlSrcSetCandidate> candidates,
        string? sizes, string? defaultSource, HtmlResponsiveImageSelectionOptions options) {
        if (candidates == null) throw new ArgumentNullException(nameof(candidates));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        bool hasWidth = false;
        bool hasDensity = false;
        var normalized = new List<NormalizedCandidate>(candidates.Count + 1);
        foreach (HtmlSrcSetCandidate candidate in candidates) {
            if (!TryReadDescriptor(candidate.Descriptor, out DescriptorKind kind, out double value)) continue;
            hasWidth |= kind == DescriptorKind.Width;
            hasDensity |= kind != DescriptorKind.Width;
            normalized.Add(new NormalizedCandidate(candidate, kind, value, UsesDefaultSource: false));
        }
        if (hasWidth && hasDensity) return default;
        if (!hasWidth && !string.IsNullOrWhiteSpace(defaultSource)
            && !normalized.Any(candidate => candidate.Value == 1D)) {
            normalized.Add(new NormalizedCandidate(new HtmlSrcSetCandidate(defaultSource!, string.Empty), DescriptorKind.Density, 1D,
                UsesDefaultSource: true));
        }
        if (normalized.Count == 0) return default;

        double sourceSize = hasWidth ? ResolveSourceSize(sizes, options) : options.ViewportWidth;
        var densities = new List<(HtmlSrcSetCandidate Candidate, double Density, bool UsesDefaultSource)>(normalized.Count);
        var seen = new HashSet<double>();
        foreach (NormalizedCandidate candidate in normalized) {
            double density = candidate.Kind == DescriptorKind.Width ? candidate.Value / sourceSize : candidate.Value;
            if (!IsFinite(density) || density <= 0D || !seen.Add(density)) continue;
            densities.Add((candidate.Candidate, density, candidate.UsesDefaultSource));
        }
        if (densities.Count == 0) return default;
        densities.Sort((left, right) => left.Density.CompareTo(right.Density));
        (HtmlSrcSetCandidate Candidate, double Density, bool UsesDefaultSource) selected = densities[densities.Count - 1];
        foreach ((HtmlSrcSetCandidate Candidate, double Density, bool UsesDefaultSource) candidate in densities) {
            if (candidate.Density + 0.000000001D < options.DevicePixelRatio) continue;
            selected = candidate;
            break;
        }
        return new HtmlResponsiveImageSelection(selected.Candidate, selected.Density, sourceSize, selected.UsesDefaultSource);
    }

    private static double ResolveSourceSize(string? sizes, HtmlResponsiveImageSelectionOptions options) {
        if (string.IsNullOrWhiteSpace(sizes)) return options.ViewportWidth;
        if (options.MaxSizesCharacters.HasValue && sizes!.Length > options.MaxSizesCharacters.Value) return options.ViewportWidth;
        foreach (string entry in SplitTopLevel(sizes!)) {
            string item = entry.Trim();
            if (item.Length == 0 || item.Equals("auto", StringComparison.OrdinalIgnoreCase)) continue;
            if (TryFindFinalComponent(item, out int boundary)) {
                string length = item.Substring(boundary).Trim();
                if (TryResolveSourceLength(length, options, out double value)) {
                    string media = item.Substring(0, boundary).Trim();
                    if (media.Length == 0 || HtmlComputedStyleEngine.IsApplicableMedia(media, options.MediaContext,
                        options.ViewportWidth, options.ViewportHeight, options.MediaFeatures)) return value;
                }
            }
            if (TryResolveSourceLength(item, options, out double bare)) return bare;
        }
        return options.ViewportWidth;
    }

    private static bool TryResolveSourceLength(string value, HtmlResponsiveImageSelectionOptions options, out double result) {
        result = 0D;
        if (value.IndexOf('%') >= 0 || value.Equals("auto", StringComparison.OrdinalIgnoreCase)) return false;
        return HtmlRenderCssValues.TryLength(value, options.ViewportWidth, options.DefaultFontSize,
            options.DefaultFontSize, options.ViewportWidth, options.ViewportHeight, out result)
            && IsFinite(result) && result > 0D;
    }

    private static bool TryFindFinalComponent(string value, out int start) {
        int depth = 0;
        bool inWhitespace = false;
        int finalWhitespaceStart = -1;
        int finalComponentStart = -1;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '(') depth++;
            else if (current == ')' && depth > 0) depth--;
            bool whitespace = depth == 0 && char.IsWhiteSpace(current);
            if (whitespace && !inWhitespace) finalWhitespaceStart = index;
            if (!whitespace && inWhitespace) finalComponentStart = index;
            inWhitespace = whitespace;
        }
        start = finalComponentStart;
        return finalWhitespaceStart >= 0 && finalComponentStart > finalWhitespaceStart;
    }

    private static IEnumerable<string> SplitTopLevel(string value) {
        int start = 0;
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '(') depth++;
            else if (current == ')' && depth > 0) depth--;
            else if (current == ',' && depth == 0) {
                yield return value.Substring(start, index - start);
                start = index + 1;
            }
        }
        yield return value.Substring(start);
    }

    private static bool TryReadDescriptor(string descriptor, out DescriptorKind kind, out double value) {
        kind = DescriptorKind.Density;
        value = 1D;
        string normalized = (descriptor ?? string.Empty).Trim();
        if (normalized.Length == 0) return true;
        string[] parts = normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        string primary = parts.FirstOrDefault(part => part.EndsWith("w", StringComparison.Ordinal))
            ?? parts.FirstOrDefault(part => part.EndsWith("x", StringComparison.Ordinal))
            ?? string.Empty;
        if (primary.Length < 2) return false;
        if (!double.TryParse(primary.Substring(0, primary.Length - 1),
            System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value)
            || !IsFinite(value) || value <= 0D) return false;
        kind = primary.EndsWith("w", StringComparison.Ordinal) ? DescriptorKind.Width : DescriptorKind.Density;
        return true;
    }

    private enum DescriptorKind { Density, Width }
    private readonly record struct NormalizedCandidate(HtmlSrcSetCandidate Candidate, DescriptorKind Kind, double Value,
        bool UsesDefaultSource);

    internal static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
