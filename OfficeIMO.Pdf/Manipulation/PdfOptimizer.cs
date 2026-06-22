using System.IO.Compression;

namespace OfficeIMO.Pdf;

/// <summary>Applies dependency-free, lossless PDF optimization actions.</summary>
public static class PdfOptimizer {
    /// <summary>Optimizes a PDF byte array with lossless actions.</summary>
    public static PdfOptimizationActionResult Optimize(byte[] pdf, PdfOptimizationOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfOptimizationOptions effectiveOptions = (options ?? new PdfOptimizationOptions()).Clone();
        if (effectiveOptions.MinimumStreamCompressionBytes < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Minimum stream compression size cannot be negative.");
        }

        PdfDocumentProbe probe = PdfInspector.Probe(pdf);
        if (probe.Security.HasEncryption) {
            throw new NotSupportedException("Encrypted PDF files are not supported for lossless optimization by OfficeIMO.Pdf yet.");
        }

        if (probe.Security.HasXrefStreams) {
            throw new NotSupportedException("XRef stream PDFs are not supported for lossless optimization by OfficeIMO.Pdf yet.");
        }

        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        PdfOptimizationReport reportBefore = PdfDiagnostics.AnalyzeOptimization(pdf);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber <= 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        var actions = new List<PdfOptimizationAction>();
        var skippedActions = new List<PdfOptimizationSkippedAction>();
        var optimizedObjects = new Dictionary<int, PdfIndirectObject>(objects);
        if (effectiveOptions.CompressUnfilteredStreams) {
            CompressUnfilteredStreams(optimizedObjects, effectiveOptions, actions, skippedActions);
        }

        byte[] candidate = RewriteAllObjects(optimizedObjects, catalogObjectNumber, pdf);
        PdfOptimizationReport reportAfter = PdfDiagnostics.AnalyzeOptimization(candidate);
        if (effectiveOptions.KeepOriginalWhenNotSmaller && candidate.Length >= pdf.Length) {
            return new PdfOptimizationActionResult(
                (byte[])pdf.Clone(),
                pdf.LongLength,
                pdf.LongLength,
                candidate.LongLength,
                reportBefore,
                reportAfter,
                actions.AsReadOnly(),
                skippedActions.AsReadOnly(),
                returnedOriginal: true);
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
            returnedOriginal: false);
    }

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

        File.WriteAllBytes(outputPath, result.Bytes);
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
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        return output.ToArray();
    }

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, byte[] sourcePdf) {
        int[] sourceIds = objects.Keys.OrderBy(static id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            rewritten.Add(PdfPageExtractor.WrapObject(numberMap[sourceId], PdfPageExtractor.SerializeObject(objects[sourceId].Value, context)));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(new PdfMetadata()))));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion);
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
