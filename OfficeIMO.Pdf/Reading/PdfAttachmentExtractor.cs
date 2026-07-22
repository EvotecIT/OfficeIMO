using OfficeIMO.Drawing.Internal;
using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>
/// Extracts embedded file attachments from PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static class PdfAttachmentExtractor {
    /// <summary>Extracts only attachments accepted by the caller predicate.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(byte[] pdf, Func<PdfExtractedAttachment, bool> predicate) {
        Guard.NotNull(predicate, nameof(predicate));
        return ExtractAttachments(pdf).Where(predicate).ToArray();
    }

    /// <summary>Extracts attachments with an exact file name.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachmentsByFileName(byte[] pdf, string fileName) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName));
        return ExtractAttachments(pdf, attachment => string.Equals(attachment.FileName, fileName, StringComparison.Ordinal));
    }

    /// <summary>Extracts embedded file attachments from a PDF byte array.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        return ExtractAttachments(objects, trailerRaw);
    }

    /// <summary>Extracts embedded file attachments from a PDF file path.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractAttachments(File.ReadAllBytes(path));
    }

    /// <summary>Extracts embedded file attachments from the current position of a readable stream.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return ExtractAttachments(buffer.ToArray());
    }

    /// <summary>Extracts embedded file attachments from a parsed PDF document.</summary>
    public static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(PdfReadDocument document) {
        Guard.NotNull(document, nameof(document));
        return document.ExtractAttachments();
    }

    /// <summary>Extracts embedded file attachments from a PDF path and writes them to <paramref name="outputDirectory"/>.</summary>
    public static IReadOnlyList<string> ExtractAttachments(string inputPath, string outputDirectory) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        return WriteAttachmentFiles(ExtractAttachments(inputPath), fullOutputDirectory);
    }

    /// <summary>Extracts embedded file attachments from a byte array and writes them to <paramref name="outputDirectory"/>.</summary>
    public static IReadOnlyList<string> ExtractAttachments(byte[] pdf, string outputDirectory) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        return WriteAttachmentFiles(ExtractAttachments(pdf), fullOutputDirectory);
    }

    /// <summary>Extracts embedded file attachments from a stream and writes them to <paramref name="outputDirectory"/>.</summary>
    public static IReadOnlyList<string> ExtractAttachments(Stream stream, string outputDirectory) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        return WriteAttachmentFiles(ExtractAttachments(stream), fullOutputDirectory);
    }

    internal static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(Dictionary<int, PdfIndirectObject> objects, string trailerRaw, PdfReadLimits? limits = null) {
        Guard.NotNull(objects, nameof(objects));
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        effectiveLimits.Validate();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null) {
            return Array.Empty<PdfExtractedAttachment>();
        }

        var attachments = new List<PdfExtractedAttachment>();
        var decodedEmbeddedStreams = new Dictionary<PdfStream, byte[]>();
        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDictionary(objects, namesObject) is PdfDictionary namesDictionary &&
            namesDictionary.Items.TryGetValue("EmbeddedFiles", out var embeddedFilesTreeObject)) {
            var visitedTrees = new HashSet<int>();
            ReadEmbeddedFilesNameTree(objects, embeddedFilesTreeObject, attachments, visitedTrees, decodedEmbeddedStreams, effectiveLimits.MaxDecodedStreamBytes);
        }

        foreach (PdfArray associatedFiles in PdfAssociatedFileGraph.FindAssociatedFileArrays(objects)) {
            ReadAssociatedFiles(objects, associatedFiles, attachments, decodedEmbeddedStreams, effectiveLimits.MaxDecodedStreamBytes);
        }

        ReadFileAttachmentAnnotations(objects, attachments, decodedEmbeddedStreams, effectiveLimits.MaxDecodedStreamBytes);

        return attachments.Count == 0 ? Array.Empty<PdfExtractedAttachment>() : attachments.AsReadOnly();
    }

    private static void ReadFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        List<PdfExtractedAttachment> attachments,
        Dictionary<PdfStream, byte[]> decodedEmbeddedStreams,
        int maximumDecodedStreamBytes) {
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            ReadFileAttachmentAnnotations(objects, indirect.Value, attachments, visited, decodedEmbeddedStreams, maximumDecodedStreamBytes);
        }
    }

    private static void ReadFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        List<PdfExtractedAttachment> attachments,
        ISet<PdfObject> visited,
        Dictionary<PdfStream, byte[]> decodedEmbeddedStreams,
        int maximumDecodedStreamBytes) {
        if (!visited.Add(value)) return;
        if (value is PdfStream stream) {
            ReadFileAttachmentAnnotations(objects, stream.Dictionary, attachments, visited, decodedEmbeddedStreams, maximumDecodedStreamBytes);
            return;
        }
        if (value is PdfDictionary dictionary) {
            if (string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("FS", out PdfObject? fileSpecObject)) {
                int referencedFileSpecObjectNumber = fileSpecObject is PdfReference fileSpecReference
                    ? fileSpecReference.ObjectNumber
                    : 0;
                bool alreadyDecoded = referencedFileSpecObjectNumber > 0 &&
                    attachments.Any(attachment => attachment.FileSpecObjectNumber == referencedFileSpecObjectNumber);
                if (!alreadyDecoded) {
                    string name = TryReadFileSpecName(objects, fileSpecObject) ?? "FileAttachment";
                    PdfExtractedAttachment? attachment = TryBuildAttachment(
                        objects,
                        name,
                        fileSpecObject,
                        "FileAttachment",
                        decodedEmbeddedStreams,
                        maximumDecodedStreamBytes);
                    if (attachment != null && !ContainsAttachment(attachments, attachment)) {
                        attachments.Add(attachment);
                    }
                }
            }

            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) {
                    ReadFileAttachmentAnnotations(objects, child, attachments, visited, decodedEmbeddedStreams, maximumDecodedStreamBytes);
                }
            }
            return;
        }
        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) {
                    ReadFileAttachmentAnnotations(objects, child, attachments, visited, decodedEmbeddedStreams, maximumDecodedStreamBytes);
                }
            }
        }
    }

    private static bool ContainsAttachment(
        IEnumerable<PdfExtractedAttachment> attachments,
        PdfExtractedAttachment candidate) =>
        attachments.Any(attachment =>
            candidate.FileSpecObjectNumber > 0 && attachment.FileSpecObjectNumber == candidate.FileSpecObjectNumber ||
            candidate.EmbeddedFileObjectNumber > 0 && attachment.EmbeddedFileObjectNumber == candidate.EmbeddedFileObjectNumber);

    private static void ReadEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject treeObject,
        List<PdfExtractedAttachment> attachments,
        HashSet<int> visitedTrees,
        Dictionary<PdfStream, byte[]> decodedEmbeddedStreams,
        int maximumDecodedStreamBytes) {
        int treeObjectNumber = treeObject is PdfReference treeReference ? treeReference.ObjectNumber : 0;
        if (treeObjectNumber > 0 && !visitedTrees.Add(treeObjectNumber)) {
            return;
        }

        if (ResolveDictionary(objects, treeObject) is not PdfDictionary tree) {
            return;
        }

        if (ResolveObject(objects, tree.Items.TryGetValue("Names", out var namesObject) ? namesObject : null) is PdfArray names) {
            for (int i = 0; i + 1 < names.Items.Count; i += 2) {
                if (ResolveObject(objects, names.Items[i]) is not PdfStringObj name) {
                    continue;
                }

                PdfObject fileSpecObject = names.Items[i + 1];
                PdfExtractedAttachment? attachment = TryBuildAttachment(objects, name.Value, fileSpecObject, "Names/EmbeddedFiles", decodedEmbeddedStreams, maximumDecodedStreamBytes);
                if (attachment != null) {
                    attachments.Add(attachment);
                }
            }
        }

        if (ResolveObject(objects, tree.Items.TryGetValue("Kids", out var kidsObject) ? kidsObject : null) is PdfArray kids) {
            foreach (PdfObject kid in kids.Items) {
                ReadEmbeddedFilesNameTree(objects, kid, attachments, visitedTrees, decodedEmbeddedStreams, maximumDecodedStreamBytes);
            }
        }
    }

    private static void ReadAssociatedFiles(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray associatedFiles,
        List<PdfExtractedAttachment> attachments,
        Dictionary<PdfStream, byte[]> decodedEmbeddedStreams,
        int maximumDecodedStreamBytes) {
        var existingFileSpecs = new HashSet<int>();
        foreach (PdfExtractedAttachment attachment in attachments) {
            if (attachment.FileSpecObjectNumber > 0) {
                existingFileSpecs.Add(attachment.FileSpecObjectNumber);
            }
        }

        for (int i = 0; i < associatedFiles.Items.Count; i++) {
            PdfObject fileSpecObject = associatedFiles.Items[i];
            int fileSpecObjectNumber = fileSpecObject is PdfReference reference ? reference.ObjectNumber : 0;
            if (fileSpecObjectNumber > 0 && existingFileSpecs.Contains(fileSpecObjectNumber)) {
                continue;
            }

            string name = TryReadFileSpecName(objects, fileSpecObject) ?? "AF." + i.ToString(CultureInfo.InvariantCulture);
            PdfExtractedAttachment? attachment = TryBuildAttachment(objects, name, fileSpecObject, "AF", decodedEmbeddedStreams, maximumDecodedStreamBytes);
            if (attachment != null) {
                attachments.Add(attachment);
                if (attachment.FileSpecObjectNumber > 0) {
                    existingFileSpecs.Add(attachment.FileSpecObjectNumber);
                }
            }
        }
    }

    private static PdfExtractedAttachment? TryBuildAttachment(
        Dictionary<int, PdfIndirectObject> objects,
        string name,
        PdfObject fileSpecObject,
        string source,
        Dictionary<PdfStream, byte[]> decodedEmbeddedStreams,
        int maximumDecodedStreamBytes) {
        int fileSpecObjectNumber = fileSpecObject is PdfReference fileSpecReference ? fileSpecReference.ObjectNumber : 0;
        if (ResolveDictionary(objects, fileSpecObject) is not PdfDictionary fileSpec ||
            ResolveDictionary(objects, fileSpec.Items.TryGetValue("EF", out var embeddedFilesObject) ? embeddedFilesObject : null) is not PdfDictionary embeddedFiles) {
            return null;
        }

        PdfObject? embeddedFileObject = embeddedFiles.Items.TryGetValue("UF", out var unicodeEmbeddedFileObject)
            ? unicodeEmbeddedFileObject
            : embeddedFiles.Items.TryGetValue("F", out var regularEmbeddedFileObject)
                ? regularEmbeddedFileObject
                : null;

        int embeddedFileObjectNumber = embeddedFileObject is PdfReference embeddedFileReference ? embeddedFileReference.ObjectNumber : 0;
        if (ResolveObject(objects, embeddedFileObject) is not PdfStream stream) {
            return null;
        }

        string fileName = TryReadText(objects, fileSpec, "F") ?? name;
        string? unicodeFileName = TryReadText(objects, fileSpec, "UF");
        string? description = TryReadText(objects, fileSpec, "Desc");
        string? mimeType = TryReadStreamSubtype(objects, stream.Dictionary);
        PdfAssociatedFileRelationship relationship = TryReadRelationship(objects, fileSpec);
        string filter = GetFilterName(objects, stream.Dictionary.Items.TryGetValue("Filter", out var filterObject) ? filterObject : null);
        if (decodedEmbeddedStreams.ContainsKey(stream)) return null;
        byte[] bytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes);
        decodedEmbeddedStreams[stream] = bytes;
        PdfDictionary? parameters = ResolveDictionary(objects, stream.Dictionary.Items.TryGetValue("Params", out PdfObject? parametersObject) ? parametersObject : null);
        DateTimeOffset? creationDate = TryReadPdfDate(objects, parameters, "CreationDate");
        DateTimeOffset? modificationDate = TryReadPdfDate(objects, parameters, "ModDate");

        return new PdfExtractedAttachment(
            name,
            fileName,
            unicodeFileName,
            description,
            mimeType,
            relationship,
            filter,
            fileSpecObjectNumber,
            embeddedFileObjectNumber,
            bytes,
            source,
            creationDate,
            modificationDate);
    }

    private static DateTimeOffset? TryReadPdfDate(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? dictionary, string key) {
        if (dictionary == null ||
            !dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(objects, value) is not PdfStringObj text) return null;
        string raw = text.Value.StartsWith("D:", StringComparison.Ordinal) ? text.Value.Substring(2) : text.Value;
        if (raw.Length < 4 || !TryPart(raw, 0, 4, out int year)) return null;
        int month = TryPart(raw, 4, 2, out int parsedMonth) ? parsedMonth : 1;
        int day = TryPart(raw, 6, 2, out int parsedDay) ? parsedDay : 1;
        int hour = TryPart(raw, 8, 2, out int parsedHour) ? parsedHour : 0;
        int minute = TryPart(raw, 10, 2, out int parsedMinute) ? parsedMinute : 0;
        int second = TryPart(raw, 12, 2, out int parsedSecond) ? parsedSecond : 0;
        TimeSpan offset = TimeSpan.Zero;
        if (raw.Length > 14 && (raw[14] == '+' || raw[14] == '-')) {
            int offsetHour = TryPart(raw, 15, 2, out int parsedOffsetHour) ? parsedOffsetHour : 0;
            int minuteIndex = raw.Length > 17 && raw[17] == '\'' ? 18 : 17;
            int offsetMinute = TryPart(raw, minuteIndex, 2, out int parsedOffsetMinute) ? parsedOffsetMinute : 0;
            offset = new TimeSpan(offsetHour, offsetMinute, 0);
            if (raw[14] == '-') offset = -offset;
        }
        try { return new DateTimeOffset(year, month, day, hour, minute, second, offset); } catch (ArgumentOutOfRangeException) { return null; }
    }

    private static bool TryPart(string value, int index, int length, out int result) {
        result = 0;
        if (index < 0 || index + length > value.Length) return false;
        for (int i = 0; i < length; i++) {
            char character = value[index + i];
            if (character < '0' || character > '9') { result = 0; return false; }
            result = (result * 10) + (character - '0');
        }
        return true;
    }

    private static string? TryReadFileSpecName(Dictionary<int, PdfIndirectObject> objects, PdfObject fileSpecObject) {
        if (ResolveDictionary(objects, fileSpecObject) is not PdfDictionary fileSpec) {
            return null;
        }

        return TryReadText(objects, fileSpec, "UF") ?? TryReadText(objects, fileSpec, "F");
    }

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfStringObj text &&
            text.Value.Length > 0
            ? text.Value
            : null;
    }

    private static string? TryReadStreamSubtype(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        return ResolveObject(objects, dictionary.Items.TryGetValue("Subtype", out var subtypeObject) ? subtypeObject : null) is PdfName subtype &&
            subtype.Name.Length > 0
            ? subtype.Name
            : null;
    }

    private static PdfAssociatedFileRelationship TryReadRelationship(Dictionary<int, PdfIndirectObject> objects, PdfDictionary fileSpec) {
        if (ResolveObject(objects, fileSpec.Items.TryGetValue("AFRelationship", out var relationshipObject) ? relationshipObject : null) is not PdfName relationship) {
            return PdfAssociatedFileRelationship.Unspecified;
        }

        return relationship.Name switch {
            "Source" => PdfAssociatedFileRelationship.Source,
            "Data" => PdfAssociatedFileRelationship.Data,
            "Alternative" => PdfAssociatedFileRelationship.Alternative,
            "Supplement" => PdfAssociatedFileRelationship.Supplement,
            _ => PdfAssociatedFileRelationship.Unspecified
        };
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? obj) {
        return ResolveObject(objects, obj) as PdfDictionary;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? obj) {
        return PdfObjectLookup.Resolve(objects, obj);
    }

    private static string GetFilterName(Dictionary<int, PdfIndirectObject> objects, PdfObject? obj) {
        PdfObject? resolved = ResolveObject(objects, obj);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array) {
            var names = new List<string>();
            foreach (PdfObject item in array.Items) {
                if (ResolveObject(objects, item) is PdfName itemName) {
                    names.Add(itemName.Name);
                }
            }

            return string.Join(",", names);
        }

        return string.Empty;
    }

    private static List<string> WriteAttachmentFiles(IReadOnlyList<PdfExtractedAttachment> attachments, string outputDirectory) {
        var paths = new List<string>(attachments.Count);
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < attachments.Count; i++) {
            PdfExtractedAttachment attachment = attachments[i];
            string fileName = GetSafeFileName(attachment.UnicodeFileName ?? attachment.FileName, "attachment-" + (i + 1).ToString("0000", CultureInfo.InvariantCulture) + ".bin");
            string uniqueFileName = MakeUniqueFileName(fileName, usedNames);
            string outputPath = Path.Combine(outputDirectory, uniqueFileName);
            OfficeFileCommit.WriteAllBytes(outputPath, attachment.Bytes);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static string GetSafeFileName(string? fileName, string fallback) {
        string safe = Path.GetFileName(fileName ?? string.Empty);
        if (string.IsNullOrWhiteSpace(safe)) {
            safe = fallback;
        }

        char[] invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(safe.Length);
        for (int i = 0; i < safe.Length; i++) {
            char ch = safe[i];
            sb.Append(Array.IndexOf(invalid, ch) >= 0 || char.IsControl(ch) ? '_' : ch);
        }

        safe = sb.ToString().Trim();
        return safe.Length == 0 ? fallback : safe;
    }

    private static string MakeUniqueFileName(string fileName, HashSet<string> usedNames) {
        if (usedNames.Add(fileName)) {
            return fileName;
        }

        string stem = Path.GetFileNameWithoutExtension(fileName);
        string extension = Path.GetExtension(fileName);
        for (int i = 2; ; i++) {
            string candidate = stem + "-" + i.ToString(CultureInfo.InvariantCulture) + extension;
            if (usedNames.Add(candidate)) {
                return candidate;
            }
        }
    }

    private static string ValidateOutputDirectory(string outputDirectory) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("Output directory cannot be empty or whitespace.", nameof(outputDirectory));
        }

        string fullOutputDirectory;
        try {
            fullOutputDirectory = Path.GetFullPath(outputDirectory);
        } catch (Exception ex) {
            throw new ArgumentException("Output directory is invalid.", nameof(outputDirectory), ex);
        }

        if (File.Exists(fullOutputDirectory)) {
            throw new ArgumentException("Output directory refers to a file; a directory path is required.", nameof(outputDirectory));
        }

        Directory.CreateDirectory(fullOutputDirectory);
        return fullOutputDirectory;
    }
}
