using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>
/// Extracts embedded file attachments from PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
public static class PdfAttachmentExtractor {
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

    internal static IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(Dictionary<int, PdfIndirectObject> objects, string trailerRaw) {
        Guard.NotNull(objects, nameof(objects));
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null ||
            !catalog.Items.TryGetValue("Names", out var namesObject) ||
            ResolveDictionary(objects, namesObject) is not PdfDictionary namesDictionary ||
            !namesDictionary.Items.TryGetValue("EmbeddedFiles", out var embeddedFilesTreeObject)) {
            return Array.Empty<PdfExtractedAttachment>();
        }

        var attachments = new List<PdfExtractedAttachment>();
        var visitedTrees = new HashSet<int>();
        ReadEmbeddedFilesNameTree(objects, embeddedFilesTreeObject, attachments, visitedTrees);
        return attachments.Count == 0 ? Array.Empty<PdfExtractedAttachment>() : attachments.AsReadOnly();
    }

    private static void ReadEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject treeObject,
        List<PdfExtractedAttachment> attachments,
        HashSet<int> visitedTrees) {
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
                PdfExtractedAttachment? attachment = TryBuildAttachment(objects, name.Value, fileSpecObject);
                if (attachment != null) {
                    attachments.Add(attachment);
                }
            }
        }

        if (ResolveObject(objects, tree.Items.TryGetValue("Kids", out var kidsObject) ? kidsObject : null) is PdfArray kids) {
            foreach (PdfObject kid in kids.Items) {
                ReadEmbeddedFilesNameTree(objects, kid, attachments, visitedTrees);
            }
        }
    }

    private static PdfExtractedAttachment? TryBuildAttachment(Dictionary<int, PdfIndirectObject> objects, string name, PdfObject fileSpecObject) {
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
        byte[] bytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);

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
            bytes);
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
            File.WriteAllBytes(outputPath, attachment.Bytes);
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
