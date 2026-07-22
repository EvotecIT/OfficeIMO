using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Adds, replaces, renames, and removes embedded and associated files in existing PDFs.</summary>
internal static class PdfAttachmentEditor {
    /// <summary>Applies a collection edit through the shared full-rewrite planner and validates attachment readback.</summary>
    public static PdfAttachmentEditResult Edit(byte[] pdf, Action<PdfAttachmentEditSession> edit, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(edit, nameof(edit));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAttachments, readOptions);
        IReadOnlyList<PdfExtractedAttachment> existing = PdfAttachmentExtractor.ExtractAttachments(PdfReadDocument.Open(pdf, readOptions));
        var session = new PdfAttachmentEditSession(existing.Select(static attachment =>
            new PdfAttachmentEditSource(
                new PdfEmbeddedFile(attachment.FileName, attachment.Bytes, attachment.MimeType,
                    attachment.Relationship, attachment.Description, attachment.CreationDate,
                    attachment.ModificationDate),
                attachment.FileSpecObjectNumber,
                attachment.EmbeddedFileObjectNumber)));
        edit(session);
        IReadOnlyList<PdfAttachmentEditEntry> targetEntries = session.SnapshotEntries();
        PdfEmbeddedFile[] target = targetEntries
            .Select(static entry => entry.Attachment)
            .ToArray();
        IReadOnlyDictionary<string, string> retainedOriginalNames = session.RetainedOriginalNames;
        bool removedFileAttachmentAnnotations = false;
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            removedFileAttachmentAnnotations = RewriteAttachmentGraph(
                objects, security, targetEntries, retainedOriginalNames);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
        IReadOnlyList<PdfAttachmentValidation> validations = Validate(output, target);
        ValidateAttachmentGraph(output, target.Length, readOptions);
        if (validations.Any(static validation => !validation.IsValid)) throw new InvalidOperationException("PDF attachment post-save validation failed; the artifact was not returned.");
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = readOptions,
            PreserveAnnotations = !removedFileAttachmentAnnotations,
            PreserveEmbeddedFiles = false,
            PreserveRevisionStructure = false
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(pdf, output, preservationOptions);
        return new PdfAttachmentEditResult(output, plan, preservation, validations);
    }

    /// <summary>Adds one attachment.</summary>
    public static PdfAttachmentEditResult Add(byte[] pdf, PdfEmbeddedFile attachment, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Add(attachment), readOptions);
    /// <summary>Replaces one attachment by file name.</summary>
    public static PdfAttachmentEditResult Replace(byte[] pdf, string fileName, PdfEmbeddedFile replacement, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Replace(fileName, replacement), readOptions);
    /// <summary>Renames one attachment.</summary>
    public static PdfAttachmentEditResult Rename(byte[] pdf, string fileName, string newFileName, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Rename(fileName, newFileName), readOptions);
    /// <summary>Removes one attachment.</summary>
    public static PdfAttachmentEditResult Remove(byte[] pdf, string fileName, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Remove(fileName), readOptions);

    private static bool RewriteAttachmentGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        IReadOnlyList<PdfAttachmentEditEntry> attachments,
        IReadOnlyDictionary<string, string> retainedOriginalNames) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) throw new InvalidOperationException("PDF catalog is not readable.");
        List<FileAttachmentAnnotation> fileAttachmentAnnotations = CaptureFileAttachmentAnnotations(objects);
        PdfDictionary? names = ResolveDictionary(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        names?.Items.Remove("EmbeddedFiles");
        RemoveExistingEmbeddedFileGraph(objects);
        PdfAssociatedFileGraph.RemoveAssociatedFileReferences(objects);
        if (attachments.Count == 0) {
            return ReconnectFileAttachmentAnnotations(
                objects,
                fileAttachmentAnnotations,
                retainedOriginalNames,
                new Dictionary<string, PdfReference>(StringComparer.Ordinal),
                new Dictionary<int, PdfReference>(),
                new Dictionary<int, PdfReference>());
        }

        if (names == null) { names = new PdfDictionary(); catalog.Items["Names"] = names; }
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        var nameEntries = new List<(string Name, PdfReference FileSpec)>(attachments.Count);
        var fileSpecificationsByName = new Dictionary<string, PdfReference>(StringComparer.Ordinal);
        var fileSpecificationsByOriginalObject = new Dictionary<int, PdfReference>();
        var fileSpecificationsByOriginalEmbeddedFile = new Dictionary<int, PdfReference>();
        var associated = new PdfArray();
        foreach (PdfAttachmentEditEntry entry in attachments
            .OrderBy(static item => item.Attachment.FileName, StringComparer.Ordinal)) {
            PdfEmbeddedFile attachment = entry.Attachment;
            byte[] data = attachment.DataSnapshot;
            int streamNumber = nextObjectNumber++;
            int fileSpecNumber = nextObjectNumber++;
            PdfDictionary streamDictionary = BuildEmbeddedFileDictionary(attachment, data);
            objects[streamNumber] = new PdfIndirectObject(streamNumber, 0, new PdfStream(streamDictionary, data));
            PdfDictionary fileSpec = BuildFileSpec(attachment, streamNumber);
            objects[fileSpecNumber] = new PdfIndirectObject(fileSpecNumber, 0, fileSpec);
            var reference = new PdfReference(fileSpecNumber, 0);
            nameEntries.Add((attachment.FileName, reference));
            fileSpecificationsByName[attachment.FileName] = reference;
            if (entry.SourceIdentity.FileSpecObjectNumber > 0) {
                fileSpecificationsByOriginalObject[entry.SourceIdentity.FileSpecObjectNumber] = reference;
            }
            if (entry.SourceIdentity.EmbeddedFileObjectNumber > 0) {
                fileSpecificationsByOriginalEmbeddedFile[entry.SourceIdentity.EmbeddedFileObjectNumber] = reference;
            }
            if (attachment.Relationship != PdfAssociatedFileRelationship.Unspecified) associated.Items.Add(reference);
        }
        var nameArray = new PdfArray();
        foreach ((string name, PdfReference reference) in nameEntries) { nameArray.Items.Add(new PdfStringObj(name, true)); nameArray.Items.Add(reference); }
        var embeddedFiles = new PdfDictionary(); embeddedFiles.Items["Names"] = nameArray; names.Items["EmbeddedFiles"] = embeddedFiles;
        if (associated.Items.Count > 0) catalog.Items["AF"] = associated;
        return ReconnectFileAttachmentAnnotations(
            objects,
            fileAttachmentAnnotations,
            retainedOriginalNames,
            fileSpecificationsByName,
            fileSpecificationsByOriginalObject,
            fileSpecificationsByOriginalEmbeddedFile);
    }

    private static void RemoveExistingEmbeddedFileGraph(Dictionary<int, PdfIndirectObject> objects) {
        var removedObjectNumbers = new HashSet<int>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value is PdfStream stream
                ? stream.Dictionary
                : indirect.Value as PdfDictionary;
            if (dictionary == null) continue;
            if (indirect.Value is PdfStream &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "EmbeddedFile", StringComparison.Ordinal)) {
                removedObjectNumbers.Add(indirect.ObjectNumber);
            }
            if (dictionary.Items.ContainsKey("EF")) removedObjectNumbers.Add(indirect.ObjectNumber);
        }

        foreach (PdfIndirectObject indirect in objects.Values) {
            ScrubEmbeddedFileReferences(indirect.Value, removedObjectNumbers, new HashSet<PdfObject>());
        }
        foreach (int objectNumber in removedObjectNumbers) objects.Remove(objectNumber);
    }

    private static List<FileAttachmentAnnotation> CaptureFileAttachmentAnnotations(Dictionary<int, PdfIndirectObject> objects) {
        var result = new List<FileAttachmentAnnotation>();
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            CaptureFileAttachmentAnnotations(indirect.Value, indirect.ObjectNumber, objects, result, visited);
        }
        return result;
    }

    private static void CaptureFileAttachmentAnnotations(
        PdfObject value,
        int? indirectObjectNumber,
        Dictionary<int, PdfIndirectObject> objects,
        List<FileAttachmentAnnotation> result,
        ISet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        if (value is PdfStream stream) {
            CaptureFileAttachmentAnnotations(stream.Dictionary, indirectObjectNumber, objects, result, visited);
            return;
        }
        if (value is PdfDictionary dictionary) {
            if (string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal)) {
                string? fileName = null;
                int fileSpecObjectNumber = 0;
                int embeddedFileObjectNumber = 0;
                if (dictionary.Items.TryGetValue("FS", out PdfObject? fileSpecification) &&
                    ResolveDictionary(objects, fileSpecification) is PdfDictionary fileSpecificationDictionary) {
                    fileSpecObjectNumber = fileSpecification is PdfReference fileSpecReference
                        ? fileSpecReference.ObjectNumber
                        : 0;
                    fileName = fileSpecificationDictionary.Get<PdfStringObj>("UF")?.Value
                        ?? fileSpecificationDictionary.Get<PdfStringObj>("F")?.Value;
                    PdfDictionary? embeddedFiles = ResolveDictionary(
                        objects,
                        fileSpecificationDictionary.Items.TryGetValue("EF", out PdfObject? efObject)
                            ? efObject
                            : null);
                    PdfObject? embeddedFile = embeddedFiles?.Items.TryGetValue("UF", out PdfObject? unicodeFile)
                        == true
                        ? unicodeFile
                        : embeddedFiles?.Items.TryGetValue("F", out PdfObject? regularFile) == true
                            ? regularFile
                            : null;
                    embeddedFileObjectNumber = embeddedFile is PdfReference embeddedFileReference
                        ? embeddedFileReference.ObjectNumber
                        : 0;
                }
                result.Add(new FileAttachmentAnnotation(
                    dictionary,
                    indirectObjectNumber,
                    fileName,
                    fileSpecObjectNumber,
                    embeddedFileObjectNumber));
            }
            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) CaptureFileAttachmentAnnotations(child, null, objects, result, visited);
            }
            return;
        }
        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) CaptureFileAttachmentAnnotations(child, null, objects, result, visited);
            }
        }
    }

    private static bool ReconnectFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<FileAttachmentAnnotation> annotations,
        IReadOnlyDictionary<string, string> retainedOriginalNames,
        Dictionary<string, PdfReference> fileSpecificationsByName,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalObject,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalEmbeddedFile) {
        var removedObjectNumbers = new HashSet<int>();
        var removedDirectAnnotations = new HashSet<PdfDictionary>();
        foreach (FileAttachmentAnnotation annotation in annotations) {
            PdfReference? replacement = null;
            if (annotation.FileSpecObjectNumber > 0) {
                fileSpecificationsByOriginalObject.TryGetValue(
                    annotation.FileSpecObjectNumber, out replacement);
            } else if (annotation.EmbeddedFileObjectNumber > 0) {
                fileSpecificationsByOriginalEmbeddedFile.TryGetValue(
                    annotation.EmbeddedFileObjectNumber, out replacement);
            } else if (annotation.FileName != null &&
                retainedOriginalNames.TryGetValue(annotation.FileName, out string? retainedName) &&
                fileSpecificationsByName.TryGetValue(retainedName, out PdfReference? nameReplacement)) {
                replacement = nameReplacement;
            }
            if (replacement != null) {
                annotation.Dictionary.Items["FS"] = replacement;
                continue;
            }

            if (annotation.ObjectNumber.HasValue) removedObjectNumbers.Add(annotation.ObjectNumber.Value);
            else removedDirectAnnotations.Add(annotation.Dictionary);
        }

        if (removedObjectNumbers.Count == 0 && removedDirectAnnotations.Count == 0) return false;
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            RemoveFileAttachmentAnnotationReferences(indirect.Value, removedObjectNumbers, removedDirectAnnotations, visited);
        }
        foreach (int objectNumber in removedObjectNumbers) objects.Remove(objectNumber);
        return true;
    }

    private static void RemoveFileAttachmentAnnotationReferences(
        PdfObject value,
        ISet<int> removedObjectNumbers,
        ISet<PdfDictionary> removedDirectAnnotations,
        ISet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        if (value is PdfStream stream) {
            RemoveFileAttachmentAnnotationReferences(stream.Dictionary, removedObjectNumbers, removedDirectAnnotations, visited);
            return;
        }
        if (value is PdfDictionary dictionary) {
            foreach (string key in dictionary.Items.Keys.ToArray()) {
                PdfObject child = dictionary.Items[key];
                if (child is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber) ||
                    child is PdfDictionary childDictionary && removedDirectAnnotations.Contains(childDictionary)) {
                    dictionary.Items.Remove(key);
                } else if (child is not PdfReference) {
                    RemoveFileAttachmentAnnotationReferences(child, removedObjectNumbers, removedDirectAnnotations, visited);
                }
            }
            return;
        }
        if (value is not PdfArray array) return;
        for (int index = array.Items.Count - 1; index >= 0; index--) {
            PdfObject child = array.Items[index];
            if (child is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber) ||
                child is PdfDictionary childDictionary && removedDirectAnnotations.Contains(childDictionary)) {
                array.Items.RemoveAt(index);
            } else if (child is not PdfReference) {
                RemoveFileAttachmentAnnotationReferences(child, removedObjectNumbers, removedDirectAnnotations, visited);
            }
        }
    }

    private sealed class FileAttachmentAnnotation {
        internal FileAttachmentAnnotation(
            PdfDictionary dictionary,
            int? objectNumber,
            string? fileName,
            int fileSpecObjectNumber,
            int embeddedFileObjectNumber) {
            Dictionary = dictionary;
            ObjectNumber = objectNumber;
            FileName = fileName;
            FileSpecObjectNumber = fileSpecObjectNumber;
            EmbeddedFileObjectNumber = embeddedFileObjectNumber;
        }

        internal PdfDictionary Dictionary { get; }
        internal int? ObjectNumber { get; }
        internal string? FileName { get; }
        internal int FileSpecObjectNumber { get; }
        internal int EmbeddedFileObjectNumber { get; }
    }

    private static void ScrubEmbeddedFileReferences(
        PdfObject value,
        ISet<int> removedObjectNumbers,
        ISet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        if (value is PdfStream stream) {
            ScrubEmbeddedFileReferences(stream.Dictionary, removedObjectNumbers, visited);
            return;
        }
        if (value is PdfDictionary dictionary) {
            dictionary.Items.Remove("EF");
            foreach (string key in dictionary.Items.Keys.ToArray()) {
                PdfObject child = dictionary.Items[key];
                if (child is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber)) {
                    dictionary.Items.Remove(key);
                    continue;
                }
                if (child is PdfArray array) {
                    RemoveEmbeddedFileReferences(array, removedObjectNumbers, visited);
                    if (array.Items.Count == 0 && (key == "AF" || key == "Names")) dictionary.Items.Remove(key);
                } else if (child is not PdfReference) {
                    ScrubEmbeddedFileReferences(child, removedObjectNumbers, visited);
                }
            }
            return;
        }
        if (value is PdfArray values) RemoveEmbeddedFileReferences(values, removedObjectNumbers, visited);
    }

    private static void RemoveEmbeddedFileReferences(
        PdfArray array,
        ISet<int> removedObjectNumbers,
        ISet<PdfObject> visited) {
        for (int index = array.Items.Count - 1; index >= 0; index--) {
            PdfObject child = array.Items[index];
            if (child is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber)) {
                array.Items.RemoveAt(index);
            } else if (child is not PdfReference) {
                ScrubEmbeddedFileReferences(child, removedObjectNumbers, visited);
            }
        }
    }

    private static void ValidateAttachmentGraph(byte[] pdf, int expectedCount, PdfReadOptions? readOptions) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, readOptions);
        int embeddedStreams = 0;
        int fileSpecifications = 0;
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value is PdfStream stream
                ? stream.Dictionary
                : indirect.Value as PdfDictionary;
            if (dictionary == null) continue;
            if (indirect.Value is PdfStream &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "EmbeddedFile", StringComparison.Ordinal)) {
                embeddedStreams++;
            }
            if (dictionary.Items.ContainsKey("EF")) fileSpecifications++;
        }
        if (embeddedStreams != expectedCount || fileSpecifications != expectedCount) {
            throw new InvalidOperationException("PDF attachment post-save validation found hidden or missing embedded-file objects.");
        }
    }

    private static PdfDictionary BuildEmbeddedFileDictionary(PdfEmbeddedFile attachment, byte[] data) {
        var dictionary = new PdfDictionary(); dictionary.Items["Type"] = new PdfName("EmbeddedFile"); dictionary.Items["Length"] = new PdfNumber(data.Length);
        if (!string.IsNullOrWhiteSpace(attachment.MimeType)) dictionary.Items["Subtype"] = new PdfName(attachment.MimeType!);
        var parameters = new PdfDictionary(); parameters.Items["Size"] = new PdfNumber(data.Length); parameters.Items["CheckSum"] = new PdfStringObj(ComputeChecksum(data));
        if (attachment.CreationDate.HasValue) parameters.Items["CreationDate"] = new PdfStringObj(FormatPdfDate(attachment.CreationDate.Value));
        if (attachment.ModificationDate.HasValue) parameters.Items["ModDate"] = new PdfStringObj(FormatPdfDate(attachment.ModificationDate.Value));
        dictionary.Items["Params"] = parameters;
        return dictionary;
    }

    private static PdfDictionary BuildFileSpec(PdfEmbeddedFile attachment, int streamNumber) {
        var dictionary = new PdfDictionary(); dictionary.Items["Type"] = new PdfName("Filespec"); dictionary.Items["F"] = new PdfStringObj(attachment.FileName, true); dictionary.Items["UF"] = new PdfStringObj(attachment.FileName, true);
        var embedded = new PdfDictionary(); var reference = new PdfReference(streamNumber, 0); embedded.Items["F"] = reference; embedded.Items["UF"] = reference; dictionary.Items["EF"] = embedded;
        dictionary.Items["AFRelationship"] = new PdfName(PdfEmbeddedFileDictionaryBuilder.GetRelationshipName(attachment.Relationship));
        if (attachment.Description != null) dictionary.Items["Desc"] = new PdfStringObj(attachment.Description, true);
        return dictionary;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfAttachmentValidation> Validate(byte[] pdf, PdfEmbeddedFile[] target) {
        IReadOnlyList<PdfExtractedAttachment> actual = PdfAttachmentExtractor.ExtractAttachments(pdf);
        var unmatched = actual.ToList();
        var result = new List<PdfAttachmentValidation>(target.Length);
        foreach (PdfEmbeddedFile expected in target) {
            byte[] data = expected.DataSnapshot;
            int foundIndex = unmatched.FindIndex(item =>
                string.Equals(item.FileName, expected.FileName, StringComparison.Ordinal) &&
                item.Bytes.SequenceEqual(data));
            PdfExtractedAttachment? found = foundIndex >= 0 ? unmatched[foundIndex] : null;
            if (foundIndex >= 0) unmatched.RemoveAt(foundIndex);
            bool payload = found != null && found.Bytes.SequenceEqual(data);
            bool metadata = found != null && string.Equals(found.MimeType, expected.MimeType, StringComparison.Ordinal) && string.Equals(found.Description, expected.Description, StringComparison.Ordinal) && found.Relationship == expected.Relationship && DatesMatch(found.CreationDate, expected.CreationDate) && DatesMatch(found.ModificationDate, expected.ModificationDate);
            result.Add(new PdfAttachmentValidation(expected.FileName, ToHex(ComputeChecksum(data)), payload, metadata));
        }
        if (actual.Count != target.Length) throw new InvalidOperationException("PDF attachment post-save validation found an unexpected attachment count.");
        return result.AsReadOnly();
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
    private static bool DatesMatch(DateTimeOffset? actual, DateTimeOffset? expected) => !actual.HasValue && !expected.HasValue || actual.HasValue && expected.HasValue && actual.Value.ToUnixTimeSeconds() == expected.Value.ToUnixTimeSeconds() && actual.Value.Offset == expected.Value.Offset;

    private static string FormatPdfDate(DateTimeOffset value) {
        string sign = value.Offset < TimeSpan.Zero ? "-" : "+";
        TimeSpan offset = value.Offset.Duration();
        return "D:" + value.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + sign + offset.Hours.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + "'" + offset.Minutes.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + "'";
    }

    #pragma warning disable CA5351, CA1850
    private static byte[] ComputeChecksum(byte[] data) { using (MD5 md5 = MD5.Create()) return md5.ComputeHash(data); }
    #pragma warning restore CA5351, CA1850
    private static string ToHex(byte[] bytes) { var builder = new System.Text.StringBuilder(bytes.Length * 2); for (int i = 0; i < bytes.Length; i++) builder.Append(bytes[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture)); return builder.ToString(); }
}
