using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Attributes unavailable CSS faces to original text requests before fallback removes that context.
/// Resource diagnostics are retained; only font-specific declaration failures are informational until requested.
/// </summary>
internal sealed class HtmlRenderFontFaceUsage {
    private const int MaximumCachedRequests = 4096;
    private const int MaximumCachedCharacters = 65536;
    private readonly OfficeFontFaceCollection _fonts;
    private readonly HtmlDiagnosticReport _diagnostics;
    private readonly Dictionary<string, List<UnavailableFace>> _unavailable = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<OfficeFontFace, int> _definitionOrder = new();
    private readonly HashSet<(string Families, OfficeFontFaceDescriptor Descriptor, string TextElement)> _observed = new();
    private bool _observing;
    private int _cachedCharacters;

    internal HtmlRenderFontFaceUsage(OfficeFontFaceCollection fonts, HtmlDiagnosticReport diagnostics) {
        _fonts = fonts;
        _diagnostics = diagnostics;
    }

    internal void RegisterAvailable(string family, OfficeFontFaceDescriptor descriptor, OfficeFontUnicodeRangeSet ranges, int order,
        IEnumerable<int> diagnosticIndices) {
        // Earlier rejected candidates do not degrade a face that subsequently loaded successfully.
        foreach (int index in diagnosticIndices) {
            if (CanAttribute(_diagnostics[index])) _diagnostics.Replace(index, WithImpact(_diagnostics[index], used: false));
        }
        foreach (OfficeFontFace face in _fonts.Faces) {
            if (face.Descriptor == descriptor && ReferenceEquals(face.UnicodeRanges, ranges)
                && string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)) {
                _definitionOrder[face] = order;
            }
        }
    }

    internal void RegisterUnavailable(string family, OfficeFontFaceDescriptor descriptor, OfficeFontUnicodeRangeSet ranges,
        int order, IEnumerable<int> diagnosticIndices) {
        int[] indices = diagnosticIndices.Where(index => CanAttribute(_diagnostics[index])).ToArray();
        var face = new UnavailableFace(descriptor, ranges, order, indices);
        if (!_unavailable.TryGetValue(family, out List<UnavailableFace>? faces)) {
            faces = new List<UnavailableFace>();
            _unavailable.Add(family, faces);
        }
        faces.Add(face);
        foreach (int index in indices) {
            HtmlDiagnostic diagnostic = _diagnostics[index];
            _diagnostics.Replace(index, WithImpact(diagnostic, used: false));
        }
    }

    internal void Observe(string text, string? familyNames, OfficeFontFaceDescriptor descriptor) {
        if (_unavailable.Count == 0 || _observing || string.IsNullOrEmpty(text) || string.IsNullOrWhiteSpace(familyNames)) return;
        _observing = true;
        try {
            List<string> families = OfficeFontFamilyParser.Parse(familyNames);
            if (!families.Any(family => _unavailable.ContainsKey(family))) return;
            foreach (string element in OfficeTextElements.Enumerate(text)) {
                // Match the renderer's complete grapheme requests, not independent combining scalars.
                if (string.IsNullOrWhiteSpace(element)) continue;
                long characters = (long)familyNames!.Length + element.Length;
                if (characters <= MaximumCachedCharacters) {
                    var key = (familyNames, descriptor, element);
                    if (_observed.Contains(key)) continue;
                    if (_observed.Count >= MaximumCachedRequests || _cachedCharacters + characters > MaximumCachedCharacters) {
                        _observed.Clear();
                        _cachedCharacters = 0;
                    }
                    _observed.Add(key);
                    _cachedCharacters += (int)characters;
                }
                ObserveElement(element, families, descriptor);
            }
        } finally {
            _observing = false;
        }
    }

    private void ObserveElement(string text, List<string> families, OfficeFontFaceDescriptor requested) {
        foreach (string family in families) {
            OfficeFontFace? available = _fonts.ResolveFaceInFamily(text, family, requested);
            int availableOrder = available == null ? -1
                : _definitionOrder.TryGetValue(available, out int registeredOrder) ? registeredOrder : int.MaxValue;

            if (_unavailable.TryGetValue(family, out List<UnavailableFace>? unavailable)) {
                UnavailableFace? preferred = null;
                foreach (UnavailableFace face in unavailable) {
                    if (!face.Ranges.ContainsFontCoverageText(text)) continue;
                    int rank = preferred == null ? -1 : OfficeFontFaceCollection.CompareFaceDescriptors(face.Descriptor, preferred.Descriptor, requested);
                    if (rank < 0 || rank == 0 && face.Order >= preferred!.Order) preferred = face;
                }
                if (preferred != null) {
                    int rank = available == null ? -1 : OfficeFontFaceCollection.CompareFaceDescriptors(preferred.Descriptor, available.Descriptor, requested);
                    if (rank < 0 || rank == 0 && preferred.Order > availableOrder) {
                        foreach (int index in preferred.DiagnosticIndices) {
                            _diagnostics.Replace(index, WithImpact(_diagnostics[index], used: true));
                        }
                    }
                }
            }
            // An available face in an earlier requested family makes later families irrelevant for this grapheme.
            if (available != null) return;
        }
    }

    private static bool CanAttribute(HtmlDiagnostic diagnostic) => diagnostic.Severity != HtmlDiagnosticSeverity.Error
        && diagnostic.Code is HtmlRenderDiagnosticCodes.FontFaceUnavailable or HtmlRenderDiagnosticCodes.FontFormatUnsupported
            or HtmlRenderDiagnosticCodes.FontDataUriInvalid or HtmlRenderDiagnosticCodes.ResourceContentTypeRejected or "FontResourceRejectedByPolicy";

    private static HtmlDiagnostic WithImpact(HtmlDiagnostic diagnostic, bool used) => diagnostic.WithImpact(
        used ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info,
        used ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None);

    private sealed class UnavailableFace {
        internal UnavailableFace(OfficeFontFaceDescriptor descriptor, OfficeFontUnicodeRangeSet ranges, int order, int[] diagnosticIndices) {
            Descriptor = descriptor;
            Ranges = ranges;
            Order = order;
            DiagnosticIndices = diagnosticIndices;
        }
        internal OfficeFontFaceDescriptor Descriptor { get; }
        internal OfficeFontUnicodeRangeSet Ranges { get; }
        internal int Order { get; }
        internal int[] DiagnosticIndices { get; }
    }
}
