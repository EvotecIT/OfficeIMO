using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    internal const int MaximumNamedFontFamilies = 64;

    /// <summary>
    /// Embedded named font families available to run-level text without consuming a standard-font slot.
    /// </summary>
    public IReadOnlyDictionary<string, PdfEmbeddedFontFamily> NamedFontFamilies {
        get {
            if (_namedFontFamilies == null || _namedFontFamilies.Count == 0) {
                return new ReadOnlyDictionary<string, PdfEmbeddedFontFamily>(
                    new Dictionary<string, PdfEmbeddedFontFamily>(StringComparer.OrdinalIgnoreCase));
            }

            var copy = new Dictionary<string, PdfEmbeddedFontFamily>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfEmbeddedFontFamily family in _namedFontFamilies.Values.OrderBy(value => value.FamilyName, StringComparer.OrdinalIgnoreCase)) {
                copy[family.FamilyName] = family.Clone();
            }

            return new ReadOnlyDictionary<string, PdfEmbeddedFontFamily>(copy);
        }
    }

    /// <summary>
    /// Registers a reusable embedded family under its authored family name. Named families do not
    /// consume Helvetica, Times, or Courier compatibility slots and may be used together on one page.
    /// </summary>
    public PdfOptions RegisterNamedFontFamily(PdfEmbeddedFontFamily fontFamily) {
        Guard.NotNull(fontFamily, nameof(fontFamily));
        string key = NormalizeNamedFontFamilyKey(fontFamily.FamilyName);
        if ((_namedFontFamilies?.Count ?? 0) >= MaximumNamedFontFamilies
            && _namedFontFamilies?.ContainsKey(key) != true) {
            throw new InvalidOperationException(
                $"No more than {MaximumNamedFontFamilies} named font families can be registered.");
        }
        (_namedFontFamilies ??= new Dictionary<string, PdfEmbeddedFontFamily>(StringComparer.Ordinal))[key] = fontFamily.Clone();
        RemoveNamedFontProgramCache(key);
        return this;
    }

    /// <summary>
    /// Loads and registers the first installed family in an Office/CSS family list.
    /// </summary>
    /// <param name="familyNames">Comma- or semicolon-separated installed family candidates.</param>
    /// <param name="registeredFamilyName">The registered installed family name when successful.</param>
    /// <returns>True when an installed embeddable family was found and registered.</returns>
    public bool TryRegisterNamedOfficeFontFamily(string? familyNames, out string? registeredFamilyName) {
        registeredFamilyName = null;
        if (string.IsNullOrWhiteSpace(familyNames) ||
            !TryLoadOfficeFontFamily(familyNames!, out PdfEmbeddedFontFamily? family) ||
            family == null) {
            return false;
        }

        string key = NormalizeNamedFontFamilyKey(family.FamilyName);
        if (_namedFontFamilies != null &&
            _namedFontFamilies.TryGetValue(key, out PdfEmbeddedFontFamily? registered)) {
            registeredFamilyName = registered.FamilyName;
            return true;
        }
        if ((_namedFontFamilies?.Count ?? 0) >= MaximumNamedFontFamilies) {
            return false;
        }

        RegisterNamedFontFamily(family);
        registeredFamilyName = family.FamilyName;
        return true;
    }

    /// <summary>Reports whether an embedded named family is registered.</summary>
    public bool HasNamedFontFamily(string? familyName) =>
        TryGetNamedFontFamily(familyName, out _);

    /// <summary>Removes every embedded named family and its parsed program cache.</summary>
    public PdfOptions ClearNamedFontFamilies() {
        _namedFontFamilies?.Clear();
        _namedFontPrograms?.Clear();
        _namedOpenTypeCffFontPrograms?.Clear();
        _namedFontProgramFailures?.Clear();
        return this;
    }

    internal bool TryResolveNamedFontFace(string? familyName, bool bold, bool italic, out PdfNamedFontFace face) {
        if (!TryGetNamedFontFamily(familyName, out PdfEmbeddedFontFamily? family) || family == null) {
            face = default;
            return false;
        }

        string key = NormalizeNamedFontFamilyKey(family.FamilyName);
        face = new PdfNamedFontFace(key, family.FamilyName, bold, italic);
        return true;
    }

    internal bool TryGetNamedFontProgram(PdfNamedFontFace face, out PdfTrueTypeFontProgram? fontProgram) {
        if (_namedFontPrograms != null && _namedFontPrograms.TryGetValue(face, out PdfTrueTypeFontProgram? cached)) {
            fontProgram = cached;
            return true;
        }

        if (_namedFontProgramFailures?.Contains(face) == true ||
            !TryGetNamedFontData(face, out byte[]? data, out string? fontName) ||
            data == null ||
            IsOpenTypeCffFontData(data)) {
            fontProgram = null;
            return false;
        }

        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(data, fontName);
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            AddNamedFontFailure(face, data, fontName, exception);
            (_namedFontProgramFailures ??= new HashSet<PdfNamedFontFace>()).Add(face);
            fontProgram = null;
            return false;
        }

        (_namedFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfTrueTypeFontProgram>())[face] = fontProgram;
        return true;
    }

    internal bool TryGetNamedFontProgramForGeneration(PdfNamedFontFace face, out PdfTrueTypeFontProgram? fontProgram) {
        if (!TryGetNamedFontData(face, out byte[]? data, out string? fontName) ||
            data == null ||
            IsOpenTypeCffFontData(data)) {
            fontProgram = null;
            return false;
        }

        if (_namedFontPrograms != null && _namedFontPrograms.TryGetValue(face, out PdfTrueTypeFontProgram? cached)) {
            fontProgram = cached;
            return true;
        }

        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(data, fontName);
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            AddNamedFontFailure(face, data, fontName, exception);
            throw;
        }

        (_namedFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfTrueTypeFontProgram>())[face] = fontProgram;
        _namedFontProgramFailures?.Remove(face);
        return true;
    }

    internal bool TryGetNamedOpenTypeCffFontProgram(PdfNamedFontFace face, out PdfOpenTypeCffFontProgram? fontProgram) {
        if (_namedOpenTypeCffFontPrograms != null &&
            _namedOpenTypeCffFontPrograms.TryGetValue(face, out PdfOpenTypeCffFontProgram? cached)) {
            fontProgram = cached;
            return true;
        }

        if (_namedFontProgramFailures?.Contains(face) == true ||
            !TryGetNamedFontData(face, out byte[]? data, out string? fontName) ||
            data == null ||
            !IsOpenTypeCffFontData(data)) {
            fontProgram = null;
            return false;
        }

        try {
            fontProgram = PdfOpenTypeCffFontProgram.Parse(data, fontName);
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            AddNamedFontFailure(face, data, fontName, exception);
            (_namedFontProgramFailures ??= new HashSet<PdfNamedFontFace>()).Add(face);
            fontProgram = null;
            return false;
        }

        (_namedOpenTypeCffFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfOpenTypeCffFontProgram>())[face] = fontProgram;
        return true;
    }

    internal bool TryGetNamedOpenTypeCffFontProgramForGeneration(PdfNamedFontFace face, out PdfOpenTypeCffFontProgram? fontProgram) {
        if (!TryGetNamedFontData(face, out byte[]? data, out string? fontName) ||
            data == null ||
            !IsOpenTypeCffFontData(data)) {
            fontProgram = null;
            return false;
        }

        if (_namedOpenTypeCffFontPrograms != null &&
            _namedOpenTypeCffFontPrograms.TryGetValue(face, out PdfOpenTypeCffFontProgram? cached)) {
            fontProgram = cached;
            return true;
        }

        try {
            fontProgram = PdfOpenTypeCffFontProgram.Parse(data, fontName);
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            AddNamedFontFailure(face, data, fontName, exception);
            throw;
        }

        (_namedOpenTypeCffFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfOpenTypeCffFontProgram>())[face] = fontProgram;
        _namedFontProgramFailures?.Remove(face);
        return true;
    }

    internal void ResetNamedFontProgramUsage() {
        if (_namedFontPrograms != null) {
            foreach (PdfTrueTypeFontProgram program in _namedFontPrograms.Values) {
                program.ResetGlyphUsage();
            }
        }

        if (_namedOpenTypeCffFontPrograms != null) {
            foreach (PdfOpenTypeCffFontProgram program in _namedOpenTypeCffFontPrograms.Values) {
                program.ResetGlyphUsage();
            }
        }
    }

    private bool TryGetNamedFontFamily(string? familyName, out PdfEmbeddedFontFamily? family) {
        family = null;
        if (string.IsNullOrWhiteSpace(familyName) || _namedFontFamilies == null) {
            return false;
        }

        foreach (string candidate in EnumerateOfficeFontFamilyCandidates(familyName!)) {
            if (_namedFontFamilies.TryGetValue(NormalizeNamedFontFamilyKey(candidate), out family)) {
                return true;
            }
        }

        return false;
    }

    internal bool TryGetNamedFontData(PdfNamedFontFace face, out byte[]? data, out string? fontName) {
        data = null;
        fontName = null;
        if (_namedFontFamilies == null || !_namedFontFamilies.TryGetValue(face.FamilyKey, out PdfEmbeddedFontFamily? family)) {
            return false;
        }

        string faceName;
        if (face.Bold && face.Italic) {
            data = family.BoldItalicSnapshot ?? family.BoldSnapshot ?? family.ItalicSnapshot ?? family.RegularSnapshot;
            faceName = "BoldItalic";
        } else if (face.Bold) {
            data = family.BoldSnapshot ?? family.RegularSnapshot;
            faceName = "Bold";
        } else if (face.Italic) {
            data = family.ItalicSnapshot ?? family.RegularSnapshot;
            faceName = "Italic";
        } else {
            data = family.RegularSnapshot;
            faceName = "Regular";
        }

        fontName = BuildFontFamilyFaceName(family.FamilyName, faceName);
        return true;
    }

    private void AddNamedFontFailure(PdfNamedFontFace face, byte[] data, string? fontName, Exception exception) {
        AddFontDiagnostics(
            PdfStandardFont.Helvetica,
            PdfFontDiagnostics.AnalyzeEmbeddedFontFailure(data, "named-font:" + face.FaceKey, fontName, exception));
    }

    private void RemoveNamedFontProgramCache(string familyKey) {
        if (_namedFontPrograms != null) {
            foreach (PdfNamedFontFace face in _namedFontPrograms.Keys.Where(face => face.FamilyKey == familyKey).ToArray()) {
                _namedFontPrograms.Remove(face);
            }
        }

        if (_namedOpenTypeCffFontPrograms != null) {
            foreach (PdfNamedFontFace face in _namedOpenTypeCffFontPrograms.Keys.Where(face => face.FamilyKey == familyKey).ToArray()) {
                _namedOpenTypeCffFontPrograms.Remove(face);
            }
        }

        _namedFontProgramFailures?.RemoveWhere(face => face.FamilyKey == familyKey);
    }

    private static string NormalizeNamedFontFamilyKey(string familyName) =>
        familyName.Trim().ToUpperInvariant();
}
