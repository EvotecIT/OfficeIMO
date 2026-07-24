namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxSystemFontFilesToInspect = 8192;
    internal const int MaxSystemFontDirectoriesToInspect = 512;
    internal const int MaxSystemFontDirectoryDepth = 16;
    internal const long MaxSystemFontFileBytes = 128L * 1024L * 1024L;
    internal const int MaxSystemFontFamilyCacheEntries = 32;
    private static readonly System.Collections.Generic.Dictionary<string, System.Lazy<SystemFontFamilyCacheEntry>> SystemFontFamilyCache =
        new(System.StringComparer.Ordinal);
    private static readonly object SystemFontFamilyCacheLock = new();

    /// <summary>
    /// Loads an installed TrueType font family from common operating-system font folders.
    /// </summary>
    /// <param name="familyName">Installed family name, for example <c>Arial</c>, <c>Segoe UI</c>, or <c>DejaVu Sans</c>.</param>
    /// <param name="pdfFamilyName">Optional family name to expose in generated PDF font resource names.</param>
    /// <returns>A reusable embedded font family backed by the discovered TrueType faces.</returns>
    /// <exception cref="System.IO.FileNotFoundException">No embeddable TrueType regular face was found for <paramref name="familyName"/>.</exception>
    public static PdfEmbeddedFontFamily FromSystem(string familyName, string? pdfFamilyName = null) {
        if (TryFromSystem(familyName, out PdfEmbeddedFontFamily? fontFamily, pdfFamilyName)) {
            return fontFamily!;
        }

        throw new System.IO.FileNotFoundException(
            "Could not find an embeddable TrueType font family named '" + familyName + "' in common system font folders.",
            familyName);
    }

    /// <summary>
    /// Attempts to load an installed TrueType font family from common operating-system font folders.
    /// </summary>
    /// <param name="familyName">Installed family name, for example <c>Arial</c>, <c>Segoe UI</c>, or <c>DejaVu Sans</c>.</param>
    /// <param name="fontFamily">When found, receives a reusable embedded font family backed by discovered TrueType faces.</param>
    /// <param name="pdfFamilyName">Optional family name to expose in generated PDF font resource names.</param>
    /// <returns><c>true</c> when an embeddable regular TrueType face was found; otherwise <c>false</c>.</returns>
    public static bool TryFromSystem(string familyName, out PdfEmbeddedFontFamily? fontFamily, string? pdfFamilyName = null) {
        Guard.NotNullOrWhiteSpace(familyName, nameof(familyName));
        string requestedFamily = familyName.Trim();
        string exposedFamily = string.IsNullOrWhiteSpace(pdfFamilyName) ? requestedFamily : pdfFamilyName!.Trim();
        string cacheKey = NormalizeFamilyKey(requestedFamily) + "\0" + exposedFamily;
        System.Lazy<SystemFontFamilyCacheEntry> lazyEntry;
        lock (SystemFontFamilyCacheLock) {
            if (!SystemFontFamilyCache.TryGetValue(cacheKey, out lazyEntry!)) {
                lazyEntry = new System.Lazy<SystemFontFamilyCacheEntry>(
                    () => ResolveSystemFontFamily(requestedFamily, exposedFamily),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
                if (SystemFontFamilyCache.Count < MaxSystemFontFamilyCacheEntries) {
                    SystemFontFamilyCache[cacheKey] = lazyEntry;
                }
            }
        }

        SystemFontFamilyCacheEntry entry = lazyEntry.Value;
        fontFamily = entry.FontFamily;
        return fontFamily != null;
    }

    private static SystemFontFamilyCacheEntry ResolveSystemFontFamily(string requestedFamily, string exposedFamily) {
        bool found = TryFromSystemFontFiles(
            requestedFamily,
            EnumerateSystemTrueTypeFontFiles(),
            out PdfEmbeddedFontFamily? fontFamily,
            exposedFamily);
        return new SystemFontFamilyCacheEntry(found ? fontFamily : null);
    }

    internal static bool TryFromSystemFontFiles(
        string familyName,
        System.Collections.Generic.IEnumerable<string> fontFiles,
        out PdfEmbeddedFontFamily? fontFamily,
        string? pdfFamilyName = null) {
        Guard.NotNullOrWhiteSpace(familyName, nameof(familyName));
        Guard.NotNull(fontFiles, nameof(fontFiles));
        string normalizedFamily = NormalizeFamilyKey(familyName);
        string[] acceptedFileNamePrefixes = BuildAcceptedFileNamePrefixes(normalizedFamily);

        SystemFontFaceCandidate? regularFace = null;
        SystemFontFaceCandidate? boldFace = null;
        SystemFontFaceCandidate? italicFace = null;
        SystemFontFaceCandidate? boldItalicFace = null;

        int inspectedFiles = 0;
        foreach (string fontFile in fontFiles) {
            if (inspectedFiles++ >= MaxSystemFontFilesToInspect) {
                break;
            }

            if (!TryReadSystemFontFaces(fontFile, normalizedFamily, acceptedFileNamePrefixes, out System.Collections.Generic.List<SystemFontFaceCandidate>? candidates) ||
                candidates == null) {
                continue;
            }

            for (int i = 0; i < candidates.Count; i++) {
                SystemFontFaceCandidate candidate = candidates[i];
                switch (candidate.Kind) {
                    case FontFaceKind.Regular:
                        SelectBetterFace(ref regularFace, candidate);
                        break;
                    case FontFaceKind.Bold:
                        SelectBetterFace(ref boldFace, candidate);
                        break;
                    case FontFaceKind.Italic:
                        SelectBetterFace(ref italicFace, candidate);
                        break;
                    case FontFaceKind.BoldItalic:
                        SelectBetterFace(ref boldItalicFace, candidate);
                        break;
                }
            }
        }

        if (regularFace == null) {
            fontFamily = null;
            return false;
        }

        fontFamily = new PdfEmbeddedFontFamily(
            string.IsNullOrWhiteSpace(pdfFamilyName) ? familyName : pdfFamilyName!,
            regularFace.Data,
            boldFace?.Data,
            italicFace?.Data,
            boldItalicFace?.Data);
        return true;
    }

    private static bool TryReadSystemFontFaces(string path, string normalizedMetadataFamily, string[] acceptedFileNamePrefixes, out System.Collections.Generic.List<SystemFontFaceCandidate>? candidates) {
        candidates = null;
        if (string.IsNullOrWhiteSpace(path)) {
            return false;
        }

        try {
            var fileInfo = new System.IO.FileInfo(path);
            if (!fileInfo.Exists ||
                (fileInfo.Attributes & System.IO.FileAttributes.ReparsePoint) != 0 ||
                fileInfo.Length > MaxSystemFontFileBytes) {
                return false;
            }

            byte[] fileData = System.IO.File.ReadAllBytes(path);
            System.Collections.Generic.List<byte[]> fontPrograms = ExtractTrueTypeFontPrograms(fileData);
            var found = new System.Collections.Generic.List<SystemFontFaceCandidate>();
            for (int i = 0; i < fontPrograms.Count; i++) {
                if (TryReadSystemFontFace(path, fontPrograms[i], normalizedMetadataFamily, acceptedFileNamePrefixes, out SystemFontFaceCandidate? candidate) &&
                    candidate != null) {
                    found.Add(candidate);
                }
            }

            candidates = found;
            return found.Count > 0;
        } catch (System.Exception exception) when (
            exception is System.IO.IOException ||
            exception is System.UnauthorizedAccessException ||
            exception is System.NotSupportedException ||
            exception is System.ArgumentException ||
            exception is System.ArithmeticException ||
            exception is System.FormatException ||
            exception is System.IndexOutOfRangeException ||
            exception is System.InvalidOperationException) {
            return false;
        }
    }

    private static bool TryReadSystemFontFace(string path, byte[] data, string normalizedMetadataFamily, string[] acceptedFileNamePrefixes, out SystemFontFaceCandidate? candidate) {
        candidate = null;
        try {
            _ = PdfTrueTypeFontProgram.Parse(data);
            if (TryReadTrueTypeNameMetadata(data, out TrueTypeNameMetadata? metadata) && metadata != null) {
                if (IsMetadataFamilyMatch(metadata, normalizedMetadataFamily)) {
                    FontFaceKind kind = ClassifyMetadataFace(metadata, out int metadataScore);
                    candidate = new SystemFontFaceCandidate(path, kind, metadataScore, data);
                    return true;
                }

                return false;
            }

            string fileName = System.IO.Path.GetFileNameWithoutExtension(path);
            if (!TryClassifyFace(fileName, acceptedFileNamePrefixes, out FontFaceKind faceKind)) {
                return false;
            }

            candidate = new SystemFontFaceCandidate(path, faceKind, ScoreFileNameFace(faceKind), data);
            return true;
        } catch (System.Exception exception) when (
            exception is System.IO.IOException ||
            exception is System.UnauthorizedAccessException ||
            exception is System.NotSupportedException ||
            exception is System.ArgumentException ||
            exception is System.ArithmeticException ||
            exception is System.FormatException ||
            exception is System.IndexOutOfRangeException ||
            exception is System.InvalidOperationException) {
            return false;
        }
    }

    private static void SelectBetterFace(ref SystemFontFaceCandidate? current, SystemFontFaceCandidate candidate) {
        if (current == null || candidate.Score > current.Score) {
            current = candidate;
        }
    }

    private static bool TryClassifyFace(string fileName, string[] acceptedPrefixes, out FontFaceKind faceKind) {
        string normalizedName = NormalizeFamilyKey(fileName);
        foreach (string prefix in acceptedPrefixes) {
            if (!normalizedName.StartsWith(prefix, System.StringComparison.Ordinal)) {
                continue;
            }

            string suffix = normalizedName.Substring(prefix.Length);
            faceKind = ClassifySuffix(suffix);
            return true;
        }

        faceKind = FontFaceKind.Regular;
        return false;
    }

    private static FontFaceKind ClassifySuffix(string suffix) {
        if (suffix.Length == 0 || suffix == "regular" || suffix == "r" || suffix == "mt") {
            return FontFaceKind.Regular;
        }

        if (suffix.Contains("bolditalic") || suffix.Contains("boldoblique") || suffix == "bi" || suffix == "z") {
            return FontFaceKind.BoldItalic;
        }

        if (suffix.Contains("italic") || suffix.Contains("oblique") || suffix == "i") {
            return FontFaceKind.Italic;
        }

        if (suffix.Contains("bold") || suffix == "bd" || suffix == "b") {
            return FontFaceKind.Bold;
        }

        return FontFaceKind.Regular;
    }

    private static int ScoreFileNameFace(FontFaceKind faceKind) =>
        faceKind == FontFaceKind.Regular ? 60 : 70;

    private static bool IsMetadataFamilyMatch(TrueTypeNameMetadata metadata, string normalizedMetadataFamily) {
        foreach (string? familyName in metadata.GetFamilyNames()) {
            if (string.IsNullOrWhiteSpace(familyName)) {
                continue;
            }

            if (IsMetadataFamilyNameMatch(familyName!, normalizedMetadataFamily)) {
                return true;
            }
        }

        foreach (string? faceName in metadata.GetFaceNames()) {
            if (string.IsNullOrWhiteSpace(faceName)) {
                continue;
            }

            if (IsMetadataFamilyNameMatch(faceName!, normalizedMetadataFamily)) {
                return true;
            }
        }

        return false;
    }

    internal static bool IsMetadataFamilyNameMatch(string fontFamilyName, string requestedFamilyName) =>
        string.Equals(NormalizeFamilyKey(fontFamilyName), NormalizeFamilyKey(requestedFamilyName), System.StringComparison.Ordinal);

    private static FontFaceKind ClassifyMetadataFace(TrueTypeNameMetadata metadata, out int score) {
        string primaryStyle = NormalizeFamilyKey(metadata.TypographicSubfamilyName ?? metadata.SubfamilyName ?? string.Empty);
        string fallbackStyle = NormalizeFamilyKey(
            (metadata.PostScriptName ?? string.Empty) + " " +
            (metadata.FullName ?? string.Empty));

        string style = primaryStyle.Length == 0 ? fallbackStyle : primaryStyle;
        FontFaceKind kind = ClassifyMetadataStyle(style, primaryStyle, out score);
        return kind;
    }

    private static FontFaceKind ClassifyMetadataStyle(string style, string primaryStyle, out int score) {
        bool bold = ContainsAny(style, "bold", "semibold", "demibold", "black", "heavy");
        bool italic = ContainsAny(style, "italic", "oblique");
        if (bold && italic) {
            if (string.Equals(primaryStyle, "bolditalic", System.StringComparison.Ordinal) ||
                string.Equals(primaryStyle, "boldoblique", System.StringComparison.Ordinal)) {
                score = 120;
            } else if (primaryStyle.Contains("bold") && (primaryStyle.Contains("italic") || primaryStyle.Contains("oblique"))) {
                score = primaryStyle.Contains("semibold") || primaryStyle.Contains("demibold") ? 100 : 105;
            } else {
                score = 85;
            }

            return FontFaceKind.BoldItalic;
        }

        if (italic) {
            score = primaryStyle.Contains("italic") || primaryStyle.Contains("oblique") ? 105 : 80;
            return FontFaceKind.Italic;
        }

        if (bold) {
            score = string.Equals(primaryStyle, "bold", System.StringComparison.Ordinal) ? 105 : 80;
            return FontFaceKind.Bold;
        }

        if (primaryStyle.Length == 0 ||
            string.Equals(primaryStyle, "regular", System.StringComparison.Ordinal) ||
            string.Equals(primaryStyle, "normal", System.StringComparison.Ordinal) ||
            string.Equals(primaryStyle, "book", System.StringComparison.Ordinal) ||
            string.Equals(primaryStyle, "roman", System.StringComparison.Ordinal)) {
            score = 105;
            return FontFaceKind.Regular;
        }

        score = 55;
        return FontFaceKind.Regular;
    }

    internal static int GetMetadataStyleScore(string styleName) {
        string normalizedStyle = NormalizeFamilyKey(styleName);
        _ = ClassifyMetadataStyle(normalizedStyle, normalizedStyle, out int score);
        return score;
    }

    private static bool ContainsAny(string value, params string[] needles) {
        for (int i = 0; i < needles.Length; i++) {
            if (value.Contains(needles[i])) {
                return true;
            }
        }

        return false;
    }

    private static string[] BuildAcceptedFileNamePrefixes(string normalizedFamily) {
        if (normalizedFamily == "timesnewroman") {
            return new[] { "timesnewroman", "times" };
        }

        if (normalizedFamily == "couriernew") {
            return new[] { "couriernew", "cour" };
        }

        if (normalizedFamily == "segoeui") {
            return new[] { "segoeui", "segui" };
        }

        return new[] { normalizedFamily };
    }

    private static string NormalizeFamilyKey(string value) {
        var builder = new System.Text.StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsLetterOrDigit(character)) {
                builder.Append(char.ToLowerInvariant(character));
            }
        }

        return builder.ToString();
    }

    private static System.Collections.Generic.IEnumerable<string> EnumerateSystemTrueTypeFontFiles() {
        var seenRoots = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
        foreach (string root in GetSystemFontRoots()) {
            if (root.Length == 0 || !seenRoots.Add(root) || !System.IO.Directory.Exists(root)) {
                continue;
            }

            foreach (string file in EnumerateTrueTypeFontFiles(root)) {
                yield return file;
            }
        }
    }

    internal static System.Collections.Generic.IEnumerable<string> GetSystemFontRoots() {
        string windows = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Windows);
        if (!string.IsNullOrWhiteSpace(windows)) {
            yield return System.IO.Path.Combine(windows, "Fonts");
        }

        string localAppData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
        if (!string.IsNullOrWhiteSpace(localAppData)) {
            yield return System.IO.Path.Combine(localAppData, "Microsoft", "Windows", "Fonts");
        }

        yield return "/usr/share/fonts";
        yield return "/usr/local/share/fonts";

        string userProfile = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrWhiteSpace(userProfile)) {
            yield return System.IO.Path.Combine(userProfile, ".local", "share", "fonts");
            yield return System.IO.Path.Combine(userProfile, ".fonts");
            yield return System.IO.Path.Combine(userProfile, "Library", "Fonts");
        }

        yield return "/Library/Fonts";
        yield return "/System/Library/Fonts";
    }

    internal static System.Collections.Generic.IEnumerable<string> EnumerateTrueTypeFontFiles(string root) {
        var directories = new System.Collections.Generic.Stack<(string Path, int Depth)>();
        var seenDirectories = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
        directories.Push((root, 0));
        int inspectedDirectories = 0;

        while (directories.Count > 0 && inspectedDirectories < MaxSystemFontDirectoriesToInspect) {
            (string current, int depth) = directories.Pop();
            string canonicalPath;
            try {
                canonicalPath = System.IO.Path.GetFullPath(current);
                var directoryInfo = new System.IO.DirectoryInfo(canonicalPath);
                if (!directoryInfo.Exists ||
                    (directoryInfo.Attributes & System.IO.FileAttributes.ReparsePoint) != 0 ||
                    !seenDirectories.Add(canonicalPath)) {
                    continue;
                }
            } catch (System.Exception exception) when (
                exception is System.IO.IOException ||
                exception is System.UnauthorizedAccessException ||
                exception is System.NotSupportedException ||
                exception is System.ArgumentException) {
                continue;
            }

            inspectedDirectories++;
            System.Collections.Generic.IEnumerable<string> files;
            try {
                files = System.IO.Directory.EnumerateFiles(canonicalPath);
            } catch (System.Exception exception) when (
                exception is System.IO.IOException ||
                exception is System.UnauthorizedAccessException) {
                files = System.Array.Empty<string>();
            }

            using (System.Collections.Generic.IEnumerator<string> enumerator = files.GetEnumerator()) {
                while (true) {
                    string file;
                    try {
                        if (!enumerator.MoveNext()) {
                            break;
                        }

                        file = enumerator.Current;
                    } catch (System.Exception exception) when (
                        exception is System.IO.IOException ||
                        exception is System.UnauthorizedAccessException) {
                        break;
                    }

                    string extension = System.IO.Path.GetExtension(file);
                    if (string.Equals(extension, ".ttf", System.StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(extension, ".ttc", System.StringComparison.OrdinalIgnoreCase)) {
                        yield return file;
                    }
                }
            }

            if (depth >= MaxSystemFontDirectoryDepth) {
                continue;
            }

            System.Collections.Generic.IEnumerable<string> children;
            try {
                children = System.IO.Directory.EnumerateDirectories(canonicalPath);
            } catch (System.Exception exception) when (
                exception is System.IO.IOException ||
                exception is System.UnauthorizedAccessException) {
                continue;
            }

            using (System.Collections.Generic.IEnumerator<string> enumerator = children.GetEnumerator()) {
                while (directories.Count + inspectedDirectories < MaxSystemFontDirectoriesToInspect) {
                    try {
                        if (!enumerator.MoveNext()) {
                            break;
                        }

                        directories.Push((enumerator.Current, depth + 1));
                    } catch (System.Exception exception) when (
                        exception is System.IO.IOException ||
                        exception is System.UnauthorizedAccessException) {
                        break;
                    }
                }
            }
        }
    }

    private static bool TryReadTrueTypeNameMetadata(byte[] data, out TrueTypeNameMetadata? metadata) {
        metadata = null;
        try {
            if (data.Length < 12) {
                return false;
            }

            var tables = ReadFontTableDirectory(data);
            if (!tables.TryGetValue("name", out FontTableRecord nameTable)) {
                return false;
            }

            var names = new System.Collections.Generic.Dictionary<int, TrueTypeNameValue>();
            int offset = nameTable.Offset;
            int count = ReadUInt16(data, offset + 2);
            int stringOffset = offset + ReadUInt16(data, offset + 4);
            for (int i = 0; i < count; i++) {
                int record = offset + 6 + i * 12;
                EnsureRange(data, record, 12);
                int platformId = ReadUInt16(data, record);
                int encodingId = ReadUInt16(data, record + 2);
                int nameId = ReadUInt16(data, record + 6);
                if (nameId != 1 && nameId != 2 && nameId != 4 && nameId != 6 && nameId != 16 && nameId != 17) {
                    continue;
                }

                int length = ReadUInt16(data, record + 8);
                int valueOffset = stringOffset + ReadUInt16(data, record + 10);
                EnsureRange(data, valueOffset, length);
                string? value = DecodeNameValue(data, valueOffset, length, platformId, encodingId);
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                int score = GetNameValueScore(platformId);
                if (!names.TryGetValue(nameId, out TrueTypeNameValue? existing) || score > existing.Score) {
                    names[nameId] = new TrueTypeNameValue(value!.Trim(), score);
                }
            }

            metadata = new TrueTypeNameMetadata(
                GetName(names, 1),
                GetName(names, 2),
                GetName(names, 4),
                GetName(names, 6),
                GetName(names, 16),
                GetName(names, 17));
            return true;
        } catch (System.Exception exception) when (exception is System.NotSupportedException) {
            return false;
        }
    }

    private static System.Collections.Generic.Dictionary<string, FontTableRecord> ReadFontTableDirectory(byte[] data) {
        int numTables = ReadUInt16(data, 4);
        var tables = new System.Collections.Generic.Dictionary<string, FontTableRecord>(System.StringComparer.Ordinal);
        int recordOffset = 12;
        for (int index = 0; index < numTables; index++) {
            int offset = recordOffset + index * 16;
            EnsureRange(data, offset, 16);
            string tag = System.Text.Encoding.ASCII.GetString(data, offset, 4);
            uint tableOffset = ReadUInt32(data, offset + 8);
            uint tableLength = ReadUInt32(data, offset + 12);
            if (tableOffset > int.MaxValue || tableLength > int.MaxValue) {
                throw new System.NotSupportedException("TrueType font table offsets are too large.");
            }

            EnsureRange(data, (int)tableOffset, (int)tableLength);
            tables[tag] = new FontTableRecord((int)tableOffset, (int)tableLength);
        }

        return tables;
    }

    private static string? DecodeNameValue(byte[] data, int offset, int length, int platformId, int encodingId) {
        if (platformId == 3 || platformId == 0) {
            return length % 2 == 0
                ? System.Text.Encoding.BigEndianUnicode.GetString(data, offset, length).TrimEnd('\0')
                : null;
        }

        if (platformId == 1 && encodingId == 0) {
            return System.Text.Encoding.ASCII.GetString(data, offset, length).TrimEnd('\0');
        }

        return null;
    }

    private static int GetNameValueScore(int platformId) {
        if (platformId == 3) {
            return 30;
        }

        if (platformId == 0) {
            return 20;
        }

        return 10;
    }

    private static string? GetName(System.Collections.Generic.Dictionary<int, TrueTypeNameValue> names, int nameId) =>
        names.TryGetValue(nameId, out TrueTypeNameValue? value) ? value.Value : null;

    private static ushort ReadUInt16(byte[] data, int offset) {
        EnsureRange(data, offset, 2);
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static uint ReadUInt32(byte[] data, int offset) {
        EnsureRange(data, offset, 4);
        return ((uint)data[offset] << 24) |
            ((uint)data[offset + 1] << 16) |
            ((uint)data[offset + 2] << 8) |
            data[offset + 3];
    }

    private static void EnsureRange(byte[] data, int offset, int length) {
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            throw new System.NotSupportedException("TrueType font table data is truncated or invalid.");
        }
    }

    private enum FontFaceKind {
        Regular,
        Bold,
        Italic,
        BoldItalic
    }

    private sealed class SystemFontFaceCandidate {
        public SystemFontFaceCandidate(string path, FontFaceKind kind, int score, byte[] data) {
            Path = path;
            Kind = kind;
            Score = score;
            Data = data;
        }

        public string Path { get; }

        public FontFaceKind Kind { get; }

        public int Score { get; }

        public byte[] Data { get; }
    }

    private sealed class SystemFontFamilyCacheEntry {
        public SystemFontFamilyCacheEntry(PdfEmbeddedFontFamily? fontFamily) {
            FontFamily = fontFamily;
        }

        public PdfEmbeddedFontFamily? FontFamily { get; }
    }

    private sealed class TrueTypeNameMetadata {
        public TrueTypeNameMetadata(
            string? familyName,
            string? subfamilyName,
            string? fullName,
            string? postScriptName,
            string? typographicFamilyName,
            string? typographicSubfamilyName) {
            FamilyName = familyName;
            SubfamilyName = subfamilyName;
            FullName = fullName;
            PostScriptName = postScriptName;
            TypographicFamilyName = typographicFamilyName;
            TypographicSubfamilyName = typographicSubfamilyName;
        }

        public string? FamilyName { get; }

        public string? SubfamilyName { get; }

        public string? FullName { get; }

        public string? PostScriptName { get; }

        public string? TypographicFamilyName { get; }

        public string? TypographicSubfamilyName { get; }

        public System.Collections.Generic.IEnumerable<string?> GetFamilyNames() {
            yield return TypographicFamilyName;
            yield return FamilyName;
        }

        public System.Collections.Generic.IEnumerable<string?> GetFaceNames() {
            yield return FullName;
            yield return PostScriptName;
            yield return CombineFamilyAndSubfamily(TypographicFamilyName, TypographicSubfamilyName);
            yield return CombineFamilyAndSubfamily(FamilyName, SubfamilyName);
        }

        private static string? CombineFamilyAndSubfamily(string? familyName, string? subfamilyName) {
            if (string.IsNullOrWhiteSpace(familyName)) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(subfamilyName) ||
                string.Equals(subfamilyName, "Regular", System.StringComparison.OrdinalIgnoreCase)) {
                return familyName;
            }

            return familyName + " " + subfamilyName;
        }
    }

    private sealed class TrueTypeNameValue {
        public TrueTypeNameValue(string value, int score) {
            Value = value;
            Score = score;
        }

        public string Value { get; }

        public int Score { get; }
    }

    private readonly struct FontTableRecord {
        public FontTableRecord(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        public int Offset { get; }

        public int Length { get; }
    }

}
