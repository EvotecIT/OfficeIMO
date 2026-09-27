using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Attributes unavailable CSS faces to original text requests before fallback removes that context.
/// Resource diagnostics are retained; only font-specific declaration failures are informational until requested.
/// </summary>
internal sealed class HtmlRenderFontFaceUsage {
    private const int MaximumCachedRequests = 4096;
    private readonly OfficeFontFaceCollection _fonts;
    private readonly HtmlDiagnosticReport _diagnostics;
    private readonly Dictionary<string, List<UnavailableFace>> _unavailable = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<OfficeFontFace, int> _definitionOrder = new();
    private readonly HashSet<(string Families, OfficeFontFaceDescriptor Descriptor, int Scalar)> _observed = new();
    private bool _observing;

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
            List<string>? families = null;
            for (int index = 0; index < text.Length; index++) {
                int scalar = text[index];
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                    scalar = char.ConvertToUtf32(text[index], text[++index]);
                } else if (char.IsSurrogate(text[index])) continue;
                // These controls do not select a separate font face. Whitespace stays with its adjacent run.
                if (char.IsWhiteSpace(text, Math.Min(index, text.Length - 1)) || scalar is 0x200C or 0x200D
                    || scalar >= 0xFE00 && scalar <= 0xFE0F || scalar >= 0xE0100 && scalar <= 0xE01EF) continue;
                if (_observed.Count >= MaximumCachedRequests) _observed.Clear();
                if (!_observed.Add((familyNames!, descriptor, scalar))) continue;
                families ??= OfficeFontFamilyParser.Parse(familyNames);
                ObserveScalar(scalar, families, descriptor);
            }
        } finally {
            _observing = false;
        }
    }

    private void ObserveScalar(int scalar, List<string> families, OfficeFontFaceDescriptor requested) {
        string text = char.ConvertFromUtf32(scalar);
        foreach (string family in families) {
            OfficeFontFace? available = null;
            int availableOrder = -1;
            foreach (OfficeFontFace face in _fonts.Faces) {
                if (!string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase)) continue;
                bool explicitResource = string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase);
                if (!(explicitResource ? face.HasGlyphs(text) : face.Covers(text))) continue;
                int order = _definitionOrder.TryGetValue(face, out int registeredOrder) ? registeredOrder : int.MaxValue;
                int rank = available == null ? -1 : OfficeFontFaceCollection.CompareFaceDescriptors(face.Descriptor, available.Descriptor, requested);
                if (rank < 0 || rank == 0 && order >= availableOrder) {
                    available = face;
                    availableOrder = order;
                }
            }

            if (_unavailable.TryGetValue(family, out List<UnavailableFace>? unavailable)) {
                UnavailableFace? preferred = null;
                foreach (UnavailableFace face in unavailable) {
                    if (!face.Ranges.Contains(scalar)) continue;
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
            // An available face in an earlier requested family makes later families irrelevant for this scalar.
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
