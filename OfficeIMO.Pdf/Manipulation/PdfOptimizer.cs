using OfficeIMO.Drawing.Internal;
using System.IO.Compression;

namespace OfficeIMO.Pdf;

/// <summary>Applies dependency-free, lossless PDF optimization actions.</summary>
internal static partial class PdfOptimizer {
    /// <summary>Optimizes a PDF byte array with lossless actions.</summary>
    public static PdfOptimizationActionResult Optimize(byte[] pdf, PdfOptimizationOptions? options = null) =>
        Optimize(pdf, options, readOptions: null);

    internal static PdfOptimizationActionResult Optimize(
        byte[] pdf,
        PdfOptimizationOptions? options,
        PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfOptimizationOptions effectiveOptions = (options ?? new PdfOptimizationOptions()).Clone();
        if (effectiveOptions.MinimumStreamCompressionBytes < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Minimum stream compression size cannot be negative.");
        }
        if (effectiveOptions.MaximumDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum decoded image bytes must be positive.");
        if (effectiveOptions.MaximumTotalDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum total decoded image bytes must be positive.");
        if (effectiveOptions.XrefFormat != PdfOptimizationXrefFormat.ClassicTable && effectiveOptions.XrefFormat != PdfOptimizationXrefFormat.XrefStream) throw new ArgumentOutOfRangeException(nameof(options), "Unsupported optimization xref format.");
        if (effectiveOptions.UseObjectStreams) effectiveOptions.XrefFormat = PdfOptimizationXrefFormat.XrefStream;
        if (effectiveOptions.Linearize && (effectiveOptions.UseObjectStreams || effectiveOptions.XrefFormat != PdfOptimizationXrefFormat.ClassicTable)) {
            throw new NotSupportedException("OfficeIMO.Pdf linearization currently requires classic cross-reference tables without object streams.");
        }

        PdfDocumentProbe probe = PdfInspector.Probe(pdf, readOptions);
        if (probe.Security.HasEncryption) {
            throw new NotSupportedException("Encrypted PDF files are not supported for lossless optimization by OfficeIMO.Pdf yet.");
        }

        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.Optimize, readOptions);

        PdfOptimizationReport reportBefore = PdfDiagnostics.AnalyzeOptimization(pdf, readOptions);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber <= 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        var actions = new List<PdfOptimizationAction>();
        var skippedActions = new List<PdfOptimizationSkippedAction>();
        var optimizedObjects = new Dictionary<int, PdfIndirectObject>(objects);
        PdfMetadata metadata = PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata;
        if (effectiveOptions.CompressUnfilteredStreams) {
            CompressUnfilteredStreams(optimizedObjects, effectiveOptions, actions, skippedActions);
        }

        if (effectiveOptions.DeduplicateIdenticalStreams) {
            DeduplicateIdenticalStreams(optimizedObjects, actions);
        }

        if (effectiveOptions.DeduplicateImages) DeduplicateImages(optimizedObjects, effectiveOptions, actions, skippedActions);
        if (effectiveOptions.DeduplicateFonts) DeduplicateTypedDictionaries(optimizedObjects, "Font", "DeduplicateFont", actions);
        if (effectiveOptions.DeduplicateResources) DeduplicateResourceDictionaries(optimizedObjects, actions);

        if (effectiveOptions.RemoveUnreferencedObjects) {
            RemoveUnreferencedObjects(optimizedObjects, catalogObjectNumber, actions);
        }

        byte[] candidate = RewriteAllObjects(
            optimizedObjects,
            catalogObjectNumber,
            metadata,
            pdf,
            PdfIncrementalObjectWriter.ReadTrailerIdEntry(trailerRaw),
            effectiveOptions);
        if (effectiveOptions.Linearize) actions.Add(new PdfOptimizationAction("Linearize", 0, pdf.LongLength, candidate.LongLength, "Reordered the document into two cross-reference sections with page and shared-object hint tables for Fast Web View."));
        if (effectiveOptions.UseObjectStreams) actions.Add(new PdfOptimizationAction("PackObjectStreams", 0, 0, 0, "Packed eligible non-stream objects into PDF 1.5 object streams."));
        if (effectiveOptions.XrefFormat == PdfOptimizationXrefFormat.XrefStream) actions.Add(new PdfOptimizationAction("WriteXrefStream", 0, 0, 0, "Emitted a PDF 1.5 cross-reference stream."));
        PdfReadOptions candidateReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, candidate.LongLength);
        PdfOptimizationReport reportAfter = PdfDiagnostics.AnalyzeOptimization(candidate, candidateReadOptions);
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = readOptions,
            RewrittenReadOptions = candidateReadOptions,
            PreserveRevisionStructure = false,
            PreserveDocumentVersionState = !effectiveOptions.UseObjectStreams && effectiveOptions.XrefFormat == PdfOptimizationXrefFormat.ClassicTable
        };
        PdfRewritePreservationReport candidatePreservation = PdfRewritePreservation.AssertPreserved(pdf, candidate, preservationOptions);
        if (!effectiveOptions.Linearize && effectiveOptions.KeepOriginalWhenNotSmaller && candidate.Length >= pdf.Length) {
            PdfRewritePreservationReport originalPreservation = PdfRewritePreservation.AssertPreserved(pdf, pdf, preservationOptions);
            return new PdfOptimizationActionResult(
                (byte[])pdf.Clone(),
                pdf.LongLength,
                pdf.LongLength,
                candidate.LongLength,
                reportBefore,
                reportAfter,
                actions.AsReadOnly(),
                skippedActions.AsReadOnly(),
                originalPreservation,
                effectiveOptions.Profile,
                effectiveOptions.XrefFormat,
                effectiveOptions.UseObjectStreams,
                effectiveOptions.Linearize,
                returnedOriginal: true,
                readOptions: readOptions);
        }

        return new PdfOptimizationActionResult(
            candidate,
            pdf.LongLength,
            candidate.LongLength,
            candidate.LongLength,
            reportBefore,
            reportAfter,
            actions.AsReadOnly(),
            skippedActions.AsReadOnly(),
            candidatePreservation,
            effectiveOptions.Profile,
            effectiveOptions.XrefFormat,
            effectiveOptions.UseObjectStreams,
            effectiveOptions.Linearize,
            returnedOriginal: false,
            readOptions: readOptions);
    }

    /// <summary>Optimizes a PDF byte array with a named deterministic profile.</summary>
    public static PdfOptimizationActionResult Optimize(byte[] pdf, PdfOptimizationProfile profile) => Optimize(pdf, PdfOptimizationOptions.Create(profile));

    internal static PdfOptimizationActionResult Optimize(
        byte[] pdf,
        PdfOptimizationProfile profile,
        PdfReadOptions? readOptions) =>
        Optimize(pdf, PdfOptimizationOptions.Create(profile), readOptions);

    /// <summary>Optimizes a PDF file with lossless actions.</summary>
    public static PdfOptimizationActionResult Optimize(string path, PdfOptimizationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Optimize(File.ReadAllBytes(path), options);
    }

    /// <summary>Optimizes a readable PDF stream from its current position with lossless actions.</summary>
    public static PdfOptimizationActionResult Optimize(Stream stream, PdfOptimizationOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Optimize(buffer.ToArray(), options);
    }

    /// <summary>Optimizes a PDF file and writes the result to another file.</summary>
    public static PdfOptimizationActionResult Optimize(string inputPath, string outputPath, PdfOptimizationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        PdfOptimizationActionResult result = Optimize(inputPath, options);
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.WriteAllBytes(outputPath, result.Bytes);
        return result;
    }

    private static void CompressUnfilteredStreams(
        Dictionary<int, PdfIndirectObject> objects,
        PdfOptimizationOptions options,
        List<PdfOptimizationAction> actions,
        List<PdfOptimizationSkippedAction> skippedActions) {
        foreach (int objectNumber in objects.Keys.OrderBy(static key => key).ToArray()) {
            PdfIndirectObject indirect = objects[objectNumber];
            if (indirect.Value is not PdfStream stream) {
                continue;
            }

            if (stream.Data.Length < options.MinimumStreamCompressionBytes) {
                skippedActions.Add(new PdfOptimizationSkippedAction(
                    "CompressStream",
                    objectNumber,
                    stream.Data.LongLength,
                    "BelowMinimumSize",
                    "Skipped unfiltered stream because it is below the configured minimum compression size."));
                continue;
            }

            if (stream.Dictionary.Items.ContainsKey("Filter")) {
                skippedActions.Add(new PdfOptimizationSkippedAction(
                    "CompressStream",
                    objectNumber,
                    stream.Data.LongLength,
                    "AlreadyFiltered",
                    "Skipped stream because it already declares a filter."));
                continue;
            }

            byte[] compressed = CompressFlate(stream.Data);
            if (compressed.Length >= stream.Data.Length) {
                skippedActions.Add(new PdfOptimizationSkippedAction(
                    "CompressStream",
                    objectNumber,
                    stream.Data.LongLength,
                    "NotSmaller",
                    "Skipped unfiltered stream because FlateDecode output was not smaller."));
                continue;
            }

            PdfDictionary dictionary = CloneStreamDictionaryForFlate(stream.Dictionary);
            objects[objectNumber] = new PdfIndirectObject(
                indirect.ObjectNumber,
                indirect.Generation,
                new PdfStream(dictionary, compressed));
            actions.Add(new PdfOptimizationAction(
                "CompressStream",
                objectNumber,
                stream.Data.LongLength,
                compressed.LongLength,
                "Compressed unfiltered stream with FlateDecode."));
        }
    }

    private static PdfDictionary CloneStreamDictionaryForFlate(PdfDictionary source) {
        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "Filter", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "DecodeParms", StringComparison.Ordinal)) {
                continue;
            }

            dictionary.Items[entry.Key] = entry.Value;
        }

        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        return dictionary;
    }

    private static byte[] CompressFlate(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        output.WriteByte((byte)((adler >> 24) & 0xFF));
        output.WriteByte((byte)((adler >> 16) & 0xFF));
        output.WriteByte((byte)((adler >> 8) & 0xFF));
        output.WriteByte((byte)(adler & 0xFF));
        return output.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static void DeduplicateIdenticalStreams(
        Dictionary<int, PdfIndirectObject> objects,
        List<PdfOptimizationAction> actions) {
        var numberMap = objects.Keys.ToDictionary(static id => id, static id => id);
        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var streamGroups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.OrderBy(static item => item.Key)) {
            if (entry.Value.Value is not PdfStream) {
                continue;
            }

            byte[] serialized = PdfPageExtractor.SerializeObject(entry.Value.Value, context);
            string fingerprint = Convert.ToBase64String(serialized);
            if (!streamGroups.TryGetValue(fingerprint, out List<int>? group)) {
                group = new List<int>();
                streamGroups[fingerprint] = group;
            }

            group.Add(entry.Key);
        }

        var replacements = new Dictionary<int, PdfReference>();
        foreach (List<int> group in streamGroups.Values) {
            if (group.Count < 2) {
                continue;
            }

            int keeper = group[0];
            int keeperGeneration = objects.TryGetValue(keeper, out PdfIndirectObject? keeperObject) ? keeperObject.Generation : 0;
            for (int i = 1; i < group.Count; i++) {
                replacements[group[i]] = new PdfReference(keeper, keeperGeneration);
            }
        }

        if (replacements.Count == 0) {
            return;
        }

        foreach (int objectNumber in objects.Keys.OrderBy(static key => key).ToArray()) {
            PdfIndirectObject indirect = objects[objectNumber];
            PdfObject rewritten = ReplaceReferences(indirect.Value, replacements);
            if (!ReferenceEquals(rewritten, indirect.Value)) {
                objects[objectNumber] = new PdfIndirectObject(indirect.ObjectNumber, indirect.Generation, rewritten);
            }
        }

        foreach (KeyValuePair<int, PdfReference> replacement in replacements.OrderBy(static item => item.Key)) {
            if (!objects.TryGetValue(replacement.Key, out PdfIndirectObject? duplicate)) {
                continue;
            }

            long originalLength = EstimateObjectLength(duplicate);
            objects.Remove(replacement.Key);
            actions.Add(new PdfOptimizationAction(
                "DeduplicateStream",
                replacement.Key,
                originalLength,
                0,
                "Rewrote references to duplicate stream object " + replacement.Key.ToString(System.Globalization.CultureInfo.InvariantCulture) + " to use object " + replacement.Value.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "."));
        }
    }

    private static PdfObject ReplaceReferences(PdfObject value, IReadOnlyDictionary<int, PdfReference> replacements) {
        if (value is PdfReference reference &&
            replacements.TryGetValue(reference.ObjectNumber, out PdfReference? replacementReference) &&
            replacementReference is not null) {
            return replacementReference;
        }

        if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                array.Items[i] = ReplaceReferences(array.Items[i], replacements);
            }

            return value;
        }

        if (value is PdfDictionary dictionary) {
            ReplaceReferences(dictionary, replacements);
            return value;
        }

        if (value is PdfStream stream) {
            ReplaceReferences(stream.Dictionary, replacements);
            return value;
        }

        return value;
    }

    private static void ReplaceReferences(PdfDictionary dictionary, IReadOnlyDictionary<int, PdfReference> replacements) {
        foreach (string key in dictionary.Items.Keys.ToArray()) {
            dictionary.Items[key] = ReplaceReferences(dictionary.Items[key], replacements);
        }
    }

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf, string trailerIdEntry, PdfOptimizationOptions options) {
        if (options.Linearize) {
            return PdfLinearizationFileAssembler.Assemble(objects, catalogObjectNumber, metadata, sourcePdf, trailerIdEntry);
        }

        int[] sourceIds = objects.Keys.OrderBy(static id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var bodies = new List<byte[]>(sourceIds.Length + 1);
        var objectStreamEligibility = new List<bool>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            bodies.Add(PdfPageExtractor.SerializeObject(objects[sourceId].Value, context));
            objectStreamEligibility.Add(sourceId != catalogObjectNumber && objects[sourceId].Value is not PdfStream);
        }

        int infoId = bodies.Count + 1;
        bodies.Add(PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata)));
        objectStreamEligibility.Add(false);

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        return PdfOptimizationFileAssembler.Assemble(bodies, objectStreamEligibility, numberMap[catalogObjectNumber], infoId, fileVersion, options, trailerIdEntry);
    }

    private static void RemoveUnreferencedObjects(
        Dictionary<int, PdfIndirectObject> objects,
        int catalogObjectNumber,
        List<PdfOptimizationAction> actions) {
        var reachable = new HashSet<int>();
        CollectReachableObjectNumbers(objects, new PdfReference(catalogObjectNumber, objects[catalogObjectNumber].Generation), reachable);
        foreach (int objectNumber in objects.Keys.OrderBy(static key => key).ToArray()) {
            if (reachable.Contains(objectNumber)) {
                continue;
            }

            long originalLength = EstimateObjectLength(objects[objectNumber]);
            objects.Remove(objectNumber);
            actions.Add(new PdfOptimizationAction(
                "RemoveUnreferencedObject",
                objectNumber,
                originalLength,
                0,
                "Removed indirect object that is not reachable from the document catalog."));
        }
    }

    private static void CollectReachableObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<int> reachable) {
        if (value is null) return;
        var pending = new Stack<PdfObject>();
        pending.Push(value);
        while (pending.Count > 0) {
            PdfObject current = pending.Pop();
            if (current is PdfReference reference) {
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                    reachable.Add(indirect.ObjectNumber)) {
                    pending.Push(indirect.Value);
                }
                continue;
            }

            if (current is PdfArray array) {
                for (int i = array.Items.Count - 1; i >= 0; i--) pending.Push(array.Items[i]);
                continue;
            }

            PdfDictionary? dictionary = current is PdfStream stream
                ? stream.Dictionary
                : current as PdfDictionary;
            if (dictionary != null) {
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            }
        }
    }

    private static long EstimateObjectLength(PdfIndirectObject indirect) {
        if (indirect.Value is PdfStream stream) {
            return stream.Data.LongLength;
        }

        return 0;
    }

    private static int FindCatalogObjectNumber(Dictionary<int, PdfIndirectObject> objects, string trailerRaw) {
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null) {
            return 0;
        }

        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (ReferenceEquals(entry.Value.Value, catalog)) {
                return entry.Key;
            }
        }

        return 0;
    }
}
