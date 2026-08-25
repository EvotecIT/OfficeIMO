using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Scoped caller-supplied font faces used by drawing measurement, rasterization, and SVG export.
/// Direct TrueType/OpenType, WOFF 1, WOFF 2, CFF/CFF2, and variable-font containers are handled by
/// the first-party engine. An optional provider can override program loading for specialized engines.
/// </summary>
public sealed class OfficeFontFaceCollection {
    private readonly List<OfficeFontFace> _faces = new List<OfficeFontFace>();
    private readonly ReadOnlyCollection<OfficeFontFace> _facesView;
    private readonly List<string> _fallbackFamilies = new List<string>();
    private readonly ReadOnlyCollection<string> _fallbackFamiliesView;

    /// <summary>Creates an empty scoped font collection.</summary>
    public OfficeFontFaceCollection() {
        _facesView = new ReadOnlyCollection<OfficeFontFace>(_faces);
        _fallbackFamiliesView = new ReadOnlyCollection<string>(_fallbackFamilies);
    }

    /// <summary>
    /// Optional decoder and outline engine for formats outside the dependency-free core.
    /// The provider reference is retained by collection clones; decoded face programs are immutable.
    /// </summary>
    public IOfficeFontProgramProvider? FontProgramProvider { get; set; }

    /// <summary>
    /// Optional variable-font axis resolver invoked for each accepted face. Values are clamped to
    /// the font-defined axis range; unknown axes and non-finite values fail closed.
    /// </summary>
    public Func<OfficeFontProgramLoadRequest, IReadOnlyDictionary<string, float>?>? FontVariationResolver { get; set; }

    /// <summary>Registered faces in registration order.</summary>
    public IReadOnlyList<OfficeFontFace> Faces => _facesView;

    /// <summary>
    /// Registered fallback families considered after the families requested by a text run.
    /// </summary>
    public IReadOnlyList<string> FallbackFamilies => _fallbackFamiliesView;

    /// <summary>Adds or replaces one family/style face. Invalid or unsupported font bytes throw.</summary>
    public OfficeFontFaceCollection Add(string familyName, byte[] data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (!TryAdd(familyName, data, style)) {
            throw new ArgumentException("The supplied bytes are not a supported TrueType outline font container.", nameof(data));
        }

        return this;
    }

    /// <summary>Attempts to add or replace one family/style face without throwing for unsupported font data.</summary>
    public bool TryAdd(string? familyName, byte[]? data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        return TryAdd(familyName, data, style, OfficeFontUnicodeRangeSet.All);
    }

    /// <summary>
    /// Adds or replaces one unicode-range-constrained family/style face.
    /// A deterministic internal resource family is assigned when the range does not cover all Unicode scalars.
    /// </summary>
    public OfficeFontFaceCollection Add(
        string familyName,
        byte[] data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet unicodeRanges) {
        if (!TryAdd(familyName, data, style, unicodeRanges)) {
            throw new ArgumentException("The supplied bytes are not a supported TrueType outline font container.", nameof(data));
        }

        return this;
    }

    /// <summary>
    /// Adds another family name for all faces already registered under an existing family.
    /// The alias reuses the immutable decoded font programs without decoding or copying them again.
    /// </summary>
    /// <param name="aliasFamilyName">Additional family name accepted by measurement and export.</param>
    /// <param name="sourceFamilyName">Existing family whose faces and unicode ranges are reused.</param>
    /// <returns>This collection for fluent configuration.</returns>
    public OfficeFontFaceCollection AddAlias(string aliasFamilyName, string sourceFamilyName) {
        if (string.IsNullOrWhiteSpace(aliasFamilyName)) {
            throw new ArgumentException("A font alias family name is required.", nameof(aliasFamilyName));
        }
        if (string.IsNullOrWhiteSpace(sourceFamilyName)) {
            throw new ArgumentException("A source font family name is required.", nameof(sourceFamilyName));
        }

        string normalizedAlias = aliasFamilyName.Trim();
        string normalizedSource = sourceFamilyName.Trim();
        var sourceFaces = new List<OfficeFontFace>();
        foreach (OfficeFontFace face in _faces) {
            if (string.Equals(face.FamilyName, normalizedSource, StringComparison.OrdinalIgnoreCase)) {
                sourceFaces.Add(face);
            }
        }
        if (sourceFaces.Count == 0) {
            throw new ArgumentException(
                $"The source font family '{normalizedSource}' is not registered.",
                nameof(sourceFamilyName));
        }
        if (string.Equals(normalizedAlias, normalizedSource, StringComparison.OrdinalIgnoreCase)) {
            return this;
        }

        for (int index = _faces.Count - 1; index >= 0; index--) {
            if (string.Equals(_faces[index].FamilyName, normalizedAlias, StringComparison.OrdinalIgnoreCase)) {
                _faces.RemoveAt(index);
            }
        }
        foreach (OfficeFontFace sourceFace in sourceFaces) {
            string resourceFamilyName = sourceFace.UnicodeRanges.IsAll
                ? normalizedAlias
                : CreateResourceFamilyName(normalizedAlias, sourceFace.Style, sourceFace.UnicodeRanges);
            _faces.Add(sourceFace.CreateAlias(normalizedAlias, resourceFamilyName));
        }

        return this;
    }

    /// <summary>
    /// Appends one registered family to the fallback chain used when requested faces do not cover a text element.
    /// </summary>
    /// <param name="familyName">Registered family or resource family name to append.</param>
    /// <returns>This collection for fluent configuration.</returns>
    public OfficeFontFaceCollection AddFallbackFamily(string familyName) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            throw new ArgumentException("A fallback font family name is required.", nameof(familyName));
        }

        string normalizedFamily = familyName.Trim();
        bool registered = false;
        foreach (OfficeFontFace face in _faces) {
            if (MatchesFamily(face, normalizedFamily)) {
                registered = true;
                break;
            }
        }
        if (!registered) {
            throw new ArgumentException(
                $"The fallback font family '{normalizedFamily}' is not registered.",
                nameof(familyName));
        }

        foreach (string existing in _fallbackFamilies) {
            if (string.Equals(existing, normalizedFamily, StringComparison.OrdinalIgnoreCase)) {
                return this;
            }
        }
        _fallbackFamilies.Add(normalizedFamily);
        return this;
    }

    /// <summary>
    /// Attempts to add or replace one unicode-range-constrained family/style face.
    /// A deterministic internal resource family is assigned when the range does not cover all Unicode scalars.
    /// </summary>
    public bool TryAdd(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet? unicodeRanges) {
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges ?? OfficeFontUnicodeRangeSet.All;
        string? resourceFamilyName = normalizedRanges.IsAll || string.IsNullOrWhiteSpace(familyName)
            ? familyName
            : CreateResourceFamilyName(familyName!.Trim(), style, normalizedRanges);
        return TryAddCore(
            familyName,
            data,
            style,
            normalizedRanges,
            resourceFamilyName,
            maximumDecodedBytes: null,
            out _,
            out _);
    }

    internal bool TryAddBounded(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet? unicodeRanges,
        int maximumDecodedBytes,
        out int decodedBytes,
        out string? error) {
        if (maximumDecodedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumDecodedBytes));
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges ?? OfficeFontUnicodeRangeSet.All;
        string? resourceFamilyName = normalizedRanges.IsAll || string.IsNullOrWhiteSpace(familyName)
            ? familyName
            : CreateResourceFamilyName(familyName!.Trim(), style, normalizedRanges);
        return TryAddCore(
            familyName,
            data,
            style,
            normalizedRanges,
            resourceFamilyName,
            maximumDecodedBytes,
            out decodedBytes,
            out error);
    }

    private bool TryAddCore(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet unicodeRanges,
        string? resourceFamilyName,
        int? maximumDecodedBytes,
        out int decodedBytes,
        out string? error) {
        decodedBytes = 0;
        error = null;
        if (string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0) {
            error = "Font data and family name are required.";
            return false;
        }

        OfficeFontContainerFormat sourceFormat = OfficeFontContainerDecoder.Detect(data);
        byte[] openTypeData;
        bool decoded = maximumDecodedBytes.HasValue
            ? OfficeFontContainerDecoder.TryDecodeToOpenType(
                data,
                maximumDecodedBytes.Value,
                out openTypeData,
                out _,
                out error)
            : OfficeFontContainerDecoder.TryDecodeToOpenType(
                data,
                out openTypeData,
                out _,
                out error);
        IReadOnlyDictionary<string, float>? variationValues = null;
        if (decoded && FontVariationResolver != null) {
            try {
                variationValues = FontVariationResolver(new OfficeFontProgramLoadRequest(
                    familyName!.Trim(),
                    openTypeData,
                    OfficeFontFace.NormalizeStyle(style),
                    OfficeFontContainerFormat.OpenType,
                    maximumDecodedBytes ?? OfficeFontContainerDecoder.DefaultMaximumDecodedBytes));
            } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                error = "The variable-font axis resolver failed: " + exception.Message;
                return false;
            }
        }
        bool isFontCollection = decoded
            && openTypeData.Length >= 4
            && openTypeData[0] == (byte)'t'
            && openTypeData[1] == (byte)'t'
            && openTypeData[2] == (byte)'c'
            && openTypeData[3] == (byte)'f';
        if (isFontCollection && variationValues != null && variationValues.Count > 0) {
            error = "Variable-font axes cannot be selected on a font collection. Extract and register the intended face as an individual OpenType font.";
            return false;
        }
        IOfficeFontProgram? builtInProgram = null;
        bool builtInVariable = false;
        if (decoded) {
            OfficeOpenTypeCffFont? cffProgram = OfficeOpenTypeCffFont.TryLoad(openTypeData, variationValues, out string? cffError);
            if (cffProgram != null) {
                builtInProgram = cffProgram;
                builtInVariable = cffProgram.IsVariable;
            } else {
                OfficeOpenTypeReader? openTypeReader = OfficeOpenTypeReader.TryCreate(openTypeData);
                OfficeFontVariationModel variationModel;
                try {
                    variationModel = openTypeReader == null
                        ? OfficeFontVariationModel.None
                        : OfficeFontVariationModel.Create(openTypeReader, variationValues);
                } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                    error = "The variable-font configuration is invalid: " + exception.Message;
                    return false;
                }
                builtInVariable = variationModel.IsVariable;
                string? trueTypeError = null;
                builtInProgram = variationModel.IsVariable
                    ? OfficeTrueTypeFont.TryLoad(openTypeData, variationModel, out trueTypeError)
                    : OfficeTrueTypeFont.TryLoad(openTypeData);
                if (builtInProgram == null && variationModel.IsVariable && !string.IsNullOrWhiteSpace(trueTypeError)) {
                    error = trueTypeError;
                }
                if (builtInProgram == null && !string.IsNullOrWhiteSpace(cffError)) error = cffError;
                if (builtInProgram == null && isFontCollection) {
                    error = "This font collection cannot be registered directly. Extract and register the intended face as an individual OpenType font.";
                }
            }
        }
        IOfficeFontProgram? parsed = builtInProgram;
        byte[] acceptedData = openTypeData;
        bool canEmbedAsStaticPdfFont = parsed != null && !builtInVariable;
        // Configuring a provider is an explicit request to use its complete layout engine even
        // for TrueType faces the dependency-free core can decode. A provider may still decline,
        // in which case the already validated built-in program remains the fallback.
        bool providerPreferred = decoded;
        if ((parsed == null || providerPreferred) && FontProgramProvider != null) {
            int providerLimit = maximumDecodedBytes ?? OfficeFontContainerDecoder.DefaultMaximumDecodedBytes;
            IReadOnlyDictionary<string, float>? providerVariationValues =
                (builtInProgram as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping
                ?? variationValues;
            // The request's container must describe the bytes handed to the provider. WOFF inputs
            // are normalized by the core before this point. Keeping the original WOFF label with
            // sfnt bytes makes a provider interpret the table directory as a web-font header.
            OfficeFontContainerFormat providerInputFormat = decoded
                ? OfficeFontContainerDecoder.Detect(openTypeData)
                : sourceFormat;
            OfficeFontProgramLoadResult? providerResult;
            try {
                providerResult = FontProgramProvider.TryLoad(new OfficeFontProgramLoadRequest(
                    familyName!.Trim(),
                    decoded ? openTypeData : data,
                    OfficeFontFace.NormalizeStyle(style),
                    providerInputFormat,
                    providerLimit,
                    providerVariationValues));
            } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                error = "The configured font-program provider failed: " + exception.Message;
                return false;
            }
            if (providerResult != null) {
                byte[]? staticData = providerResult.StaticOpenTypeDataSnapshot;
                int faceDataBytes = staticData?.Length ?? data.Length;
                long retainedBytes = (long)providerResult.DecodedByteCount + faceDataBytes;
                if (retainedBytes > providerLimit) {
                    error = "Decoded font data exceeds the configured byte limit.";
                    return false;
                }
                parsed = providerResult.Program;
                acceptedData = staticData ?? (byte[])data.Clone();
                decodedBytes = checked((int)retainedBytes);
                canEmbedAsStaticPdfFont = staticData != null;
            } else {
                parsed = builtInProgram;
                acceptedData = openTypeData;
                canEmbedAsStaticPdfFont = builtInProgram != null && !builtInVariable;
            }
        }
        if (parsed == null) {
            if (decoded && string.IsNullOrWhiteSpace(error)) error = "Decoded font data does not contain a supported outline program.";
            return false;
        }

        string normalizedFamily = familyName!.Trim();
        string normalizedResourceFamily = string.IsNullOrWhiteSpace(resourceFamilyName)
            ? normalizedFamily
            : resourceFamilyName!.Trim();
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges;
        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace existing = _faces[index];
            if (existing.Style == normalizedStyle
                && string.Equals(existing.FamilyName, normalizedFamily, StringComparison.OrdinalIgnoreCase)
                && string.Equals(existing.ResourceFamilyName, normalizedResourceFamily, StringComparison.OrdinalIgnoreCase)) {
                _faces[index] = new OfficeFontFace(
                    normalizedFamily,
                    normalizedResourceFamily,
                    acceptedData,
                    normalizedStyle,
                    normalizedRanges,
                    parsed,
                    sourceFormat,
                    canEmbedAsStaticPdfFont);
                if (decodedBytes == 0) decodedBytes = acceptedData.Length;
                return true;
            }
        }

        _faces.Add(new OfficeFontFace(
            normalizedFamily,
            normalizedResourceFamily,
            acceptedData,
            normalizedStyle,
            normalizedRanges,
            parsed,
            sourceFormat,
            canEmbedAsStaticPdfFont));
        if (decodedBytes == 0) decodedBytes = acceptedData.Length;
        return true;
    }

    /// <summary>Adds independent copies of all faces from another collection.</summary>
    public OfficeFontFaceCollection AddRange(OfficeFontFaceCollection? fonts) {
        if (fonts == null || ReferenceEquals(fonts, this)) {
            return this;
        }

        foreach (OfficeFontFace face in fonts.Faces) {
            OfficeFontFace copy = face.Clone();
            bool replaced = false;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace existing = _faces[index];
                if (existing.Style == copy.Style
                    && string.Equals(existing.FamilyName, copy.FamilyName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(existing.ResourceFamilyName, copy.ResourceFamilyName, StringComparison.OrdinalIgnoreCase)) {
                    _faces[index] = copy;
                    replaced = true;
                    break;
                }
            }
            if (!replaced) {
                _faces.Add(copy);
            }
        }
        foreach (string fallbackFamily in fonts.FallbackFamilies) {
            bool exists = false;
            foreach (string existing in _fallbackFamilies) {
                if (!string.Equals(existing, fallbackFamily, StringComparison.OrdinalIgnoreCase)) continue;
                exists = true;
                break;
            }
            if (!exists) _fallbackFamilies.Add(fallbackFamily);
        }

        return this;
    }

    /// <summary>
    /// Adds independent copies of faces from another collection without replacing faces already
    /// registered under the same family, resource family, and style. Added faces are considered
    /// before existing fallback families, while existing document-scoped faces retain precedence.
    /// </summary>
    public OfficeFontFaceCollection AddRangePreservingExisting(OfficeFontFaceCollection? fonts) {
        if (fonts == null || ReferenceEquals(fonts, this)) {
            return this;
        }

        var additions = new List<OfficeFontFace>();
        foreach (OfficeFontFace face in fonts.Faces) {
            bool exists = _faces.Exists(existing =>
                existing.Style == face.Style
                && string.Equals(existing.FamilyName, face.FamilyName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(existing.ResourceFamilyName, face.ResourceFamilyName, StringComparison.OrdinalIgnoreCase));
            if (!exists) {
                additions.Add(face.Clone());
            }
        }
        _faces.InsertRange(0, additions);

        foreach (string fallbackFamily in fonts.FallbackFamilies) {
            if (!_fallbackFamilies.Exists(existing =>
                    string.Equals(existing, fallbackFamily, StringComparison.OrdinalIgnoreCase))) {
                _fallbackFamilies.Add(fallbackFamily);
            }
        }

        return this;
    }

    /// <summary>Creates an independent collection snapshot.</summary>
    public OfficeFontFaceCollection Clone() {
        var clone = new OfficeFontFaceCollection {
            FontProgramProvider = FontProgramProvider,
            FontVariationResolver = FontVariationResolver
        };
        foreach (OfficeFontFace face in _faces) {
            clone._faces.Add(face.Clone());
        }
        clone._fallbackFamilies.AddRange(_fallbackFamilies);

        return clone;
    }

    /// <summary>Attempts to measure text with a matching scoped face.</summary>
    public bool TryMeasureText(string? text, double fontSize, string? familyNames, OfficeFontStyle style, out double width) {
        width = 0D;
        if (string.IsNullOrEmpty(text) || fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            return false;
        }
        if (_faces.Count == 0) return false;

        IReadOnlyList<OfficeFontFallbackRun> runs = PlanFallbackRuns(text, familyNames, style);
        foreach (OfficeFontFallbackRun run in runs) {
            IOfficeFontProgram? font = ResolveForText(run.Text, run.FamilyName, style, out OfficeFontStyle _);
            if (font == null) return false;
            width += font.Measure(run.Text, fontSize);
        }
        return runs.Count > 0;
    }

    internal bool TryMeasureTextElements(
        string text,
        IReadOnlyList<string> elements,
        double fontSize,
        string? familyNames,
        OfficeFontStyle style,
        out IReadOnlyList<double> widths) {
        widths = Array.Empty<double>();
        if (string.IsNullOrEmpty(text) || elements.Count == 0 || fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            return false;
        }
        if (_faces.Count == 0) return false;

        IReadOnlyList<OfficeFontFallbackRun> runs = PlanFallbackRuns(text, familyNames, style);
        var resolvedWidths = new List<double>(elements.Count);
        foreach (OfficeFontFallbackRun run in runs) {
            IOfficeFontProgram? font = ResolveForText(run.Text, run.FamilyName, style, out OfficeFontStyle _);
            if (font == null) return false;
            var runElements = new List<string>();
            foreach (string element in OfficeTextElements.Enumerate(run.Text)) runElements.Add(element);
            resolvedWidths.AddRange(font.MeasureTextElements(runElements, fontSize));
        }
        if (resolvedWidths.Count != elements.Count) return false;
        widths = resolvedWidths.AsReadOnly();
        return true;
    }

    /// <summary>
    /// Splits text into grapheme-safe runs using the first scoped family whose selected face covers each text element.
    /// Unresolved elements retain the original family list for platform or adapter fallback.
    /// </summary>
    public IReadOnlyList<OfficeFontFallbackRun> PlanFallbackRuns(string? text, string? familyNames, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (string.IsNullOrEmpty(text)) return Array.Empty<OfficeFontFallbackRun>();

        string requestedFamilies = familyNames?.Trim() ?? string.Empty;
        IReadOnlyList<OfficeFontFace> candidates = ResolveFallbackCandidates(requestedFamilies, style);
        if (candidates.Count == 0) {
            return Array.AsReadOnly(new[] { new OfficeFontFallbackRun(text!, requestedFamilies) });
        }
        var explicitlySelectedFaces = new HashSet<string>(
            OfficeFontFamilyParser.Parse(requestedFamilies),
            StringComparer.OrdinalIgnoreCase);
        var resolvedFamilies = new Dictionary<string, string>(StringComparer.Ordinal);
        var runs = new List<OfficeFontFallbackRun>();
        var currentText = new StringBuilder();
        string? currentFamily = null;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            string family;
            if (IsWhitespace(element) && currentFamily != null) {
                family = currentFamily;
            } else {
                if (!resolvedFamilies.TryGetValue(element, out string? resolvedFamily)) {
                    OfficeFontFace? face = null;
                    for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                        OfficeFontFace candidate = candidates[candidateIndex];
                        bool explicitlySelected = explicitlySelectedFaces.Contains(candidate.ResourceFamilyName);
                        if (!(explicitlySelected ? candidate.HasGlyphs(element) : candidate.Covers(element))) continue;
                        face = candidate;
                        break;
                    }
                    resolvedFamily = face?.ResourceFamilyName ?? requestedFamilies;
                    resolvedFamilies.Add(element, resolvedFamily);
                }
                family = resolvedFamily;
            }
            if (currentFamily != null && !string.Equals(currentFamily, family, StringComparison.OrdinalIgnoreCase)) {
                runs.Add(new OfficeFontFallbackRun(currentText.ToString(), currentFamily));
                currentText.Clear();
            }

            currentFamily = family;
            currentText.Append(element);
        }

        if (currentText.Length > 0) runs.Add(new OfficeFontFallbackRun(currentText.ToString(), currentFamily ?? requestedFamilies));
        return runs.AsReadOnly();
    }

    /// <summary>
    /// Resolves the scoped face that covers the supplied text for the requested CSS/Office family
    /// list and style. This uses the same unicode-range and fallback rules as measurement and
    /// rendering.
    /// </summary>
    public bool TryResolveFaceForText(
        string? text,
        string? familyNames,
        OfficeFontStyle style,
        out OfficeFontFace? face) {
        face = null;
        if (string.IsNullOrEmpty(text) || string.IsNullOrWhiteSpace(familyNames)) return false;
        ResolveForText(text!, familyNames, style, out face);
        return face != null;
    }

    private IReadOnlyList<OfficeFontFace> ResolveFallbackCandidates(string familyNames, OfficeFontStyle style) {
        if (_faces.Count == 0) return Array.Empty<OfficeFontFace>();

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        var result = new List<OfficeFontFace>();
        var added = new HashSet<OfficeFontFace>();
        var families = new List<string>();
        var addedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            if (addedFamilies.Add(family)) families.Add(family);
        }
        foreach (string family in _fallbackFamilies) {
            if (addedFamilies.Add(family)) families.Add(family);
        }
        foreach (string family in families) {
            var exact = new List<OfficeFontFace>();
            var regular = new List<OfficeFontFace>();
            var available = new List<OfficeFontFace>();
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!MatchesFamily(face, family)) continue;
                available.Add(face);
                if (face.Style == normalizedStyle) exact.Add(face);
                if (face.Style == OfficeFontStyle.Regular) regular.Add(face);
            }
            foreach (OfficeFontFace face in exact) {
                if (added.Add(face)) result.Add(face);
            }
            foreach (OfficeFontFace face in regular) {
                if (added.Add(face)) result.Add(face);
            }
            foreach (OfficeFontFace face in available) {
                if (added.Add(face)) result.Add(face);
            }
        }
        return result;
    }

    internal IOfficeFontProgram? Resolve(string? familyNames, OfficeFontStyle style) {
        return Resolve(familyNames, style, out _);
    }

    internal IOfficeFontProgram? Resolve(string? familyNames, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (string.IsNullOrEmpty(familyNames) || _faces.Count == 0) {
            return null;
        }

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            OfficeFontFace? regular = null;
            OfficeFontFace? first = null;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!MatchesFamily(face, family)) {
                    continue;
                }

                first ??= face;
                if (face.Style == normalizedStyle) {
                    resolvedStyle = face.Style;
                    return face.ParsedFont;
                }

                if (face.Style == OfficeFontStyle.Regular) {
                    regular = face;
                }
            }

            if (regular != null) {
                resolvedStyle = regular.Style;
                return regular.ParsedFont;
            }

            if (first != null) {
                resolvedStyle = first.Style;
                return first.ParsedFont;
            }
        }

        return null;
    }

    internal IOfficeFontProgram? ResolveForText(string text, string? familyNames, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        IOfficeFontProgram? font = ResolveForText(text, familyNames, style, out OfficeFontFace? face);
        resolvedStyle = face?.Style ?? OfficeFontStyle.Regular;
        return font;
    }

    private IOfficeFontProgram? ResolveForText(string text, string? familyNames, OfficeFontStyle style, out OfficeFontFace? resolvedFace) {
        resolvedFace = null;
        if (string.IsNullOrEmpty(familyNames) || _faces.Count == 0) return null;

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            OfficeFontFace? exact = null;
            OfficeFontFace? regular = null;
            OfficeFontFace? first = null;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                bool explicitlySelected = string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase);
                if (!MatchesFamily(face, family)
                    || !CoversPlannedText(face, text, requireUnicodeRange: !explicitlySelected)) continue;
                first ??= face;
                if (face.Style == normalizedStyle) exact ??= face;
                if (face.Style == OfficeFontStyle.Regular) regular ??= face;
            }

            OfficeFontFace? preferred = exact ?? regular ?? first;
            if (preferred == null) continue;
            resolvedFace = preferred;
            return preferred.ParsedFont;
        }

        return null;
    }

    private static bool IsWhitespace(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (!char.IsWhiteSpace(value[index])) return false;
        }
        return value.Length > 0;
    }

    private static bool CoversPlannedText(
        OfficeFontFace face,
        string text,
        bool requireUnicodeRange) {
        bool hasNonWhitespace = false;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            if (IsWhitespace(element)) continue;
            hasNonWhitespace = true;
            if (!(requireUnicodeRange ? face.Covers(element) : face.HasGlyphs(element))) return false;
        }
        return hasNonWhitespace || text.Length > 0;
    }

    private static bool MatchesFamily(OfficeFontFace face, string family) =>
        string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)
        || string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase);


    private static string CreateResourceFamilyName(
        string familyName,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet ranges) {
        string value = ((int)OfficeFontFace.NormalizeStyle(style)).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "|"
            + ranges.ToStableKey();
        uint hash = 2166136261;
        for (int index = 0; index < value.Length; index++) {
            hash ^= value[index];
            hash *= 16777619;
        }
        return familyName + "__officeimo_" + hash.ToString("x8", System.Globalization.CultureInfo.InvariantCulture);
    }
}
