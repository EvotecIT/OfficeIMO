namespace OfficeIMO.Pdf;

/// <summary>Dependency-free PDF diagnostics and optimization analysis.</summary>
internal static class PdfDiagnostics {
    private const long LargeStreamThresholdBytes = 1024L * 1024L;
    private const long UncompressedStreamThresholdBytes = 16L * 1024L;

    /// <summary>Analyzes a PDF byte array.</summary>
    public static PdfDiagnosticReport Analyze(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));

        PdfDocumentProbe probe = PdfInspector.Probe(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        PdfDocumentInfo? info = preflight.DocumentInfo;

        try {
            var (objects, _) = PdfSyntax.ParseObjects(pdf, options);
            return BuildReport(probe, preflight, info, objects, objectGraphError: null);
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return BuildReport(probe, preflight, info, objects: null, ex.Message);
        }
    }

    internal static PdfDiagnosticReport Analyze(
        byte[] pdf,
        PdfReadDocument document,
        PdfDocumentInfo info,
        PdfDocumentPreflight preflight) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(info, nameof(info));
        Guard.NotNull(preflight, nameof(preflight));

        return BuildReport(
            PdfInspector.Probe(pdf, document),
            preflight,
            info,
            document.Objects,
            objectGraphError: null);
    }

    /// <summary>Analyzes a PDF file.</summary>
    public static PdfDiagnosticReport Analyze(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Analyze(File.ReadAllBytes(path), options);
    }

    /// <summary>Analyzes a readable PDF stream from its current position.</summary>
    public static PdfDiagnosticReport Analyze(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Analyze(buffer.ToArray(), options);
    }

    /// <summary>Reports optimization opportunities for a PDF byte array without modifying it.</summary>
    public static PdfOptimizationReport AnalyzeOptimization(byte[] pdf, PdfReadOptions? options = null) {
        PdfDiagnosticReport diagnostics = Analyze(pdf, options);
        return BuildOptimizationReport(diagnostics);
    }

    /// <summary>Reports optimization opportunities for a PDF file without modifying it.</summary>
    public static PdfOptimizationReport AnalyzeOptimization(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return AnalyzeOptimization(File.ReadAllBytes(path), options);
    }

    /// <summary>Reports optimization opportunities for a readable PDF stream without modifying it.</summary>
    public static PdfOptimizationReport AnalyzeOptimization(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return AnalyzeOptimization(buffer.ToArray(), options);
    }

    internal static PdfOptimizationReport BuildOptimizationReport(PdfDiagnosticReport diagnostics) {
        var findings = new List<PdfDiagnosticFinding>();
        var duplicateGroups = GetDuplicateStreamGroups(diagnostics.Streams);
        long estimatedSavings = 0;
        foreach (PdfDuplicateStreamGroup group in duplicateGroups) {
            estimatedSavings += group.EstimatedSavingsBytes;
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Info,
                "DuplicateStream",
                "Duplicate stream candidate group contains " + group.ObjectNumbers.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " objects.",
                bytes: group.EstimatedSavingsBytes));
        }

        var largest = diagnostics.Streams
            .OrderByDescending(stream => stream.Length)
            .ThenBy(stream => stream.ObjectNumber)
            .Take(10)
            .ToArray();

        foreach (PdfStreamDiagnostic stream in diagnostics.Streams) {
            if (stream.Filters.Count == 0 && stream.Length >= UncompressedStreamThresholdBytes) {
                long estimate = stream.Length / 3;
                estimatedSavings += estimate;
                findings.Add(new PdfDiagnosticFinding(
                    PdfDiagnosticSeverity.Info,
                    "UncompressedStream",
                    "Stream has no filter and may benefit from compression.",
                    stream.ObjectNumber,
                    bytes: estimate));
            }

            if (string.Equals(stream.Subtype, "Image", StringComparison.Ordinal) && stream.Length >= LargeStreamThresholdBytes) {
                findings.Add(new PdfDiagnosticFinding(
                    PdfDiagnosticSeverity.Info,
                    "LargeImageStream",
                    "Large image stream may be worth reviewing for resolution or color-depth reduction.",
                    stream.ObjectNumber,
                    bytes: stream.Length));
            }
        }

        if (!diagnostics.ObjectGraphParsed && !string.IsNullOrEmpty(diagnostics.ObjectGraphError)) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "OptimizationLimited",
                "Optimization analysis is limited because the object graph could not be parsed: " + diagnostics.ObjectGraphError));
        }

        return new PdfOptimizationReport(
            diagnostics,
            duplicateGroups,
            largest,
            findings.AsReadOnly(),
            estimatedSavings);
    }

    private static void AddPreflightFindings(PdfDocumentPreflight preflight, List<PdfDiagnosticFinding> findings) {
        foreach (PdfReadBlocker blocker in preflight.ReadBlockers) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Error,
                "ReadBlocker." + blocker.Kind,
                blocker.Message));
        }

        foreach (PdfRewriteBlocker blocker in preflight.RewriteBlockers) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "RewriteBlocker." + blocker.Kind,
                blocker.Message));
        }
    }

    private static PdfDiagnosticReport BuildReport(
        PdfDocumentProbe probe,
        PdfDocumentPreflight preflight,
        PdfDocumentInfo? info,
        Dictionary<int, PdfIndirectObject>? objects,
        string? objectGraphError) {
        var findings = new List<PdfDiagnosticFinding>();
        AddPreflightFindings(preflight, findings);

        var objectTypeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var streamTypeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var fonts = new List<PdfFontDiagnostic>();
        var streams = new List<PdfStreamDiagnostic>();
        bool objectGraphParsed = objects != null;
        if (objects != null) {
            foreach (PdfIndirectObject indirect in objects.Values.OrderBy(item => item.ObjectNumber)) {
                string objectKind = GetObjectKind(indirect.Value);
                Increment(objectTypeCounts, objectKind);

                if (indirect.Value is PdfStream stream) {
                    PdfStreamDiagnostic streamDiagnostic = BuildStreamDiagnostic(indirect, stream);
                    streams.Add(streamDiagnostic);
                    Increment(streamTypeCounts, streamDiagnostic.Kind);
                    AddStreamFindings(streamDiagnostic, findings);
                } else if (indirect.Value is PdfDictionary dictionary &&
                    string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Font", StringComparison.Ordinal)) {
                    PdfFontDiagnostic font = BuildFontDiagnostic(indirect, dictionary, objects);
                    fonts.Add(font);
                    if (font.RequiresEmbeddingReview) {
                        findings.Add(new PdfDiagnosticFinding(
                            PdfDiagnosticSeverity.Warning,
                            "FontEmbeddingReview",
                            "Font dictionary does not expose an embedded font file.",
                            indirect.ObjectNumber));
                    }

                    if (font.RequiresToUnicodeReview) {
                        findings.Add(new PdfDiagnosticFinding(
                            PdfDiagnosticSeverity.Warning,
                            "FontToUnicodeReview",
                            "Composite or identity-encoded font dictionary does not expose a /ToUnicode CMap.",
                            indirect.ObjectNumber));
                    }
                }
            }
        } else {
            findings.Add(new PdfDiagnosticFinding(
                preflight.ReadBlockers.Count > 0 ? PdfDiagnosticSeverity.Warning : PdfDiagnosticSeverity.Error,
                "ObjectGraphParseFailed",
                "PDF indirect objects could not be fully parsed: " + objectGraphError));
        }

        if (probe.HasEncryption) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "EncryptionDetected",
                "Encryption markers were detected. OfficeIMO.Pdf can read Standard password-encrypted PDFs with a valid password and can perform proven authenticated page, metadata, sanitization, and simple form rewrites on unsigned inputs when the required permissions are authorized; security changes require owner authorization."));
        }

        if (probe.HasActiveContent) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "ActiveContentDetected",
                "Active content markers were detected. Review JavaScript, launch actions, and catalog actions before automated processing."));
        }

        return new PdfDiagnosticReport(
            probe,
            preflight,
            info,
            new SortedDictionary<string, int>(objectTypeCounts, StringComparer.Ordinal),
            new SortedDictionary<string, int>(streamTypeCounts, StringComparer.Ordinal),
            fonts.AsReadOnly(),
            streams.AsReadOnly(),
            findings.AsReadOnly(),
            objectGraphParsed,
            objectGraphError);
    }

    private static void AddStreamFindings(PdfStreamDiagnostic stream, List<PdfDiagnosticFinding> findings) {
        if (stream.DecodingFailed) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "StreamDecodeFailed",
                "Stream decoding failed: " + (stream.DecodingError ?? "unknown error"),
                stream.ObjectNumber,
                bytes: stream.Length));
        }

        if (stream.Length >= LargeStreamThresholdBytes) {
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Info,
                "LargeStream",
                "Large stream detected.",
                stream.ObjectNumber,
                bytes: stream.Length));
        }
    }

    private static PdfDuplicateStreamGroup[] GetDuplicateStreamGroups(IReadOnlyList<PdfStreamDiagnostic> streams) {
        var grouped = new Dictionary<string, List<PdfStreamDiagnostic>>(StringComparer.Ordinal);
        foreach (PdfStreamDiagnostic stream in streams) {
            if (stream.Length == 0) {
                continue;
            }

            string key = stream.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + stream.Hash;
            if (!grouped.TryGetValue(key, out List<PdfStreamDiagnostic>? group)) {
                group = new List<PdfStreamDiagnostic>();
                grouped.Add(key, group);
            }

            group.Add(stream);
        }

        var result = new List<PdfDuplicateStreamGroup>();
        foreach (List<PdfStreamDiagnostic> group in grouped.Values) {
            if (group.Count < 2) {
                continue;
            }

            var objectNumbers = group
                .Select(stream => stream.ObjectNumber)
                .OrderBy(number => number)
                .ToArray();
            result.Add(new PdfDuplicateStreamGroup(group[0].Hash, group[0].Length, objectNumbers));
        }

        return result
            .OrderByDescending(group => group.EstimatedSavingsBytes)
            .ThenBy(group => group.ObjectNumbers[0])
            .ToArray();
    }

    private static PdfStreamDiagnostic BuildStreamDiagnostic(PdfIndirectObject indirect, PdfStream stream) {
        string? type = stream.Dictionary.Get<PdfName>("Type")?.Name;
        string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
        string kind = !string.IsNullOrEmpty(subtype)
            ? subtype!
            : !string.IsNullOrEmpty(type) ? type! : "Stream";
        IReadOnlyList<string> filters = GetFilterNames(stream.Dictionary);
        bool decoded = filters.Count > 0 && !stream.DecodingFailed;

        return new PdfStreamDiagnostic(
            indirect.ObjectNumber,
            indirect.Generation,
            kind,
            type,
            subtype,
            filters,
            stream.Data.LongLength,
            decoded,
            stream.DecodingFailed,
            stream.DecodingError,
            GetInt(stream.Dictionary, "Width"),
            GetInt(stream.Dictionary, "Height"),
            GetInt(stream.Dictionary, "BitsPerComponent"),
            ComputeHash(stream.Data));
    }

    private static PdfFontDiagnostic BuildFontDiagnostic(PdfIndirectObject indirect, PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        string? subtype = dictionary.Get<PdfName>("Subtype")?.Name;
        string? baseFont = dictionary.Get<PdfName>("BaseFont")?.Name;
        string? encoding = dictionary.Get<PdfName>("Encoding")?.Name;
        bool hasToUnicodeMap = dictionary.Items.ContainsKey("ToUnicode");
        int? toUnicodeObjectNumber = dictionary.Items.TryGetValue("ToUnicode", out PdfObject? toUnicodeObject) &&
            toUnicodeObject is PdfReference toUnicodeReference
                ? toUnicodeReference.ObjectNumber
                : null;
        int? descriptorObjectNumber = null;
        PdfDictionary? descriptor = null;
        if (dictionary.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject)) {
            if (descriptorObject is PdfReference descriptorReference) {
                descriptorObjectNumber = descriptorReference.ObjectNumber;
            }

            descriptor = ResolveDictionary(descriptorObject, objects);
        }

        string? embeddedKind = null;
        if (descriptor is not null) {
            embeddedKind = FindEmbeddedFontFileKind(descriptor);
        }

        if (embeddedKind is null &&
            dictionary.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) &&
            ResolveObject(descendantsObject, objects) is PdfArray descendants) {
            for (int i = 0; i < descendants.Items.Count && embeddedKind is null; i++) {
                PdfDictionary? descendant = ResolveDictionary(descendants.Items[i], objects);
                if (descendant is null ||
                    !descendant.Items.TryGetValue("FontDescriptor", out PdfObject? descendantDescriptorObject)) {
                    continue;
                }

                if (!descriptorObjectNumber.HasValue && descendantDescriptorObject is PdfReference descendantDescriptorReference) {
                    descriptorObjectNumber = descendantDescriptorReference.ObjectNumber;
                }

                embeddedKind = FindEmbeddedFontFileKind(ResolveDictionary(descendantDescriptorObject, objects));
            }
        }

        return new PdfFontDiagnostic(
            indirect.ObjectNumber,
            subtype,
            baseFont,
            encoding,
            hasToUnicodeMap,
            toUnicodeObjectNumber,
            descriptorObjectNumber,
            embeddedKind is not null,
            embeddedKind);
    }

    private static string? FindEmbeddedFontFileKind(PdfDictionary? descriptor) {
        if (descriptor is null) {
            return null;
        }

        if (descriptor.Items.ContainsKey("FontFile")) {
            return "FontFile";
        }

        if (descriptor.Items.ContainsKey("FontFile2")) {
            return "FontFile2";
        }

        return descriptor.Items.ContainsKey("FontFile3") ? "FontFile3" : null;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
            return indirect.Value;
        }

        return obj;
    }

    private static PdfDictionary? ResolveDictionary(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) =>
        ResolveObject(obj, objects) as PdfDictionary;

    private static IReadOnlyList<string> GetFilterNames(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("Filter", out PdfObject? value)) {
            return Array.Empty<string>();
        }

        if (value is PdfName name) {
            return new[] { name.Name };
        }

        if (value is PdfArray array) {
            var filters = new List<string>();
            foreach (PdfObject item in array.Items) {
                if (item is PdfName filterName) {
                    filters.Add(filterName.Name);
                }
            }

            return filters.AsReadOnly();
        }

        return Array.Empty<string>();
    }

    private static int? GetInt(PdfDictionary dictionary, string key) {
        PdfNumber? number = dictionary.Get<PdfNumber>(key);
        if (number is null) {
            return null;
        }

        if (number.Value < int.MinValue || number.Value > int.MaxValue) {
            return null;
        }

        return (int)number.Value;
    }

    private static string GetObjectKind(PdfObject value) {
        if (value is PdfStream stream) {
            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (!string.IsNullOrEmpty(subtype)) {
                return "Stream." + subtype;
            }

            string? type = stream.Dictionary.Get<PdfName>("Type")?.Name;
            return string.IsNullOrEmpty(type) ? "Stream" : "Stream." + type;
        }

        if (value is PdfDictionary dictionary) {
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (!string.IsNullOrEmpty(type)) {
                return "Dictionary." + type;
            }
        }

        return value.GetType().Name;
    }

    private static void Increment(Dictionary<string, int> counts, string key) {
        if (counts.TryGetValue(key, out int count)) {
            counts[key] = count + 1;
            return;
        }

        counts.Add(key, 1);
    }

    private static string ComputeHash(byte[] data) {
        const ulong offset = 14695981039346656037UL;
        const ulong prime = 1099511628211UL;
        ulong hash = offset;
        for (int i = 0; i < data.Length; i++) {
            hash ^= data[i];
            hash *= prime;
        }

        return hash.ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }
}
