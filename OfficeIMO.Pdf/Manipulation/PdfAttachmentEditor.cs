using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Adds, replaces, renames, and removes embedded and associated files in existing PDFs.</summary>
internal static class PdfAttachmentEditor {
    /// <summary>Applies a collection edit through the shared full-rewrite planner and validates attachment readback.</summary>
    public static PdfAttachmentEditResult Edit(
        byte[] pdf,
        Action<PdfAttachmentEditSession> edit,
        PdfLoadOptions? readOptions = null,
        long? maxDecodedAttachmentBytes = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(edit, nameof(edit));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAttachments, readOptions, cancellationToken: cancellationToken);
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        IReadOnlyList<PdfExtractedAttachment> existing = maxDecodedAttachmentBytes.HasValue
            ? PdfAttachmentExtractor.ExtractAttachments(document, static _ => true, maxDecodedAttachmentBytes.Value, cancellationToken: cancellationToken)
            : document.ExtractAttachments(cancellationToken);
        if (maxDecodedAttachmentBytes.HasValue) {
            long projectedBytes = 0;
            foreach (PdfExtractedAttachment attachment in existing) {
                if (projectedBytes > maxDecodedAttachmentBytes.Value - attachment.ByteLength) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.AttachmentBytes,
                        maxDecodedAttachmentBytes.Value,
                        projectedBytes + attachment.ByteLength);
                }
                projectedBytes += attachment.ByteLength;
            }
        }
        long maximumTargetBytes = Math.Min(PdfLoadOptions.Resolve(readOptions).Limits.MaxTotalAttachmentBytes,
            maxDecodedAttachmentBytes ?? long.MaxValue);
        var session = new PdfAttachmentEditSession(existing.Select(static attachment =>
            new PdfAttachmentEditSource(
                new PdfEmbeddedFile(attachment.FileName, attachment.Bytes, attachment.MimeType,
                    attachment.Relationship, attachment.Description, attachment.CreationDate,
                    attachment.ModificationDate),
                attachment.FileSpecObjectNumber,
                attachment.EmbeddedFileObjectNumber)), maximumTargetBytes);
        cancellationToken.ThrowIfCancellationRequested();
        edit(session);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PdfAttachmentEditEntry> targetEntries = session.SnapshotEntries();
        PdfEmbeddedFile[] target = targetEntries
            .Select(static entry => entry.Attachment)
            .ToArray();
        IReadOnlyDictionary<string, string> retainedOriginalNames = session.RetainedOriginalNames;
        bool removedFileAttachmentAnnotations = false;
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            removedFileAttachmentAnnotations = RewriteAttachmentGraph(
                objects, security, targetEntries, retainedOriginalNames, cancellationToken);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        }, cancellationToken: cancellationToken);
        PdfLoadOptions outputReadOptions = PdfLoadOptions.WithMinimumInputBytes(
            PdfLoadOptions.Resolve(readOptions),
            output.LongLength);
        IReadOnlyList<PdfAttachmentValidation> validations = Validate(output, target, outputReadOptions, cancellationToken);
        ValidateAttachmentGraph(output, target.Length, outputReadOptions, cancellationToken);
        if (validations.Any(static validation => !validation.IsValid)) throw new InvalidOperationException("PDF attachment post-save validation failed; the artifact was not returned.");
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = readOptions,
            RewrittenReadOptions = outputReadOptions,
            PreserveAnnotations = !removedFileAttachmentAnnotations,
            PreserveEmbeddedFiles = false,
            PreserveRevisionStructure = false
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(pdf, output, preservationOptions, cancellationToken);
        return new PdfAttachmentEditResult(output, plan, preservation, validations, outputReadOptions);
    }

    /// <summary>Adds one attachment.</summary>
    public static PdfAttachmentEditResult Add(byte[] pdf, PdfEmbeddedFile attachment, PdfLoadOptions? readOptions = null) => Edit(pdf, session => session.Add(attachment), readOptions);
    /// <summary>Replaces one attachment by file name.</summary>
    public static PdfAttachmentEditResult Replace(byte[] pdf, string fileName, PdfEmbeddedFile replacement, PdfLoadOptions? readOptions = null) => Edit(pdf, session => session.Replace(fileName, replacement), readOptions);
    /// <summary>Renames one attachment.</summary>
    public static PdfAttachmentEditResult Rename(byte[] pdf, string fileName, string newFileName, PdfLoadOptions? readOptions = null) => Edit(pdf, session => session.Rename(fileName, newFileName), readOptions);
    /// <summary>Removes one attachment.</summary>
    public static PdfAttachmentEditResult Remove(byte[] pdf, string fileName, PdfLoadOptions? readOptions = null) => Edit(pdf, session => session.Remove(fileName), readOptions);

    private static bool RewriteAttachmentGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        IReadOnlyList<PdfAttachmentEditEntry> attachments,
        IReadOnlyDictionary<string, string> retainedOriginalNames, CancellationToken cancellationToken) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) throw new InvalidOperationException("PDF catalog is not readable.");
        List<FileAttachmentAnnotation> fileAttachmentAnnotations = CaptureFileAttachmentAnnotations(objects);
        // Structure elements, pages, annotations and XObjects can own files through /AF. Capture those owners
        // before the embedded-file graph is rebuilt so retained files stay linked to the same content.
        List<AssociatedFileOwner> associatedFileOwners = CaptureAssociatedFileOwners(objects, catalog, out AssociatedFileSet catalogAssociations, cancellationToken);
        PdfDictionary? names = ResolveDictionary(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        names?.Items.Remove("EmbeddedFiles");
        RemoveExistingEmbeddedFileGraph(objects, cancellationToken);
        PdfAssociatedFileGraph.RemoveAssociatedFileReferences(objects);
        if (attachments.Count == 0) {
            ReconnectAssociatedFileOwners(associatedFileOwners, new Dictionary<int, PdfReference>(), new Dictionary<int, PdfReference>());
            return ReconnectFileAttachmentAnnotations(
                objects,
                catalog,
                fileAttachmentAnnotations,
                retainedOriginalNames,
                new Dictionary<string, PdfReference>(StringComparer.Ordinal),
                new Dictionary<int, PdfReference>(),
                new Dictionary<int, PdfReference>(),
                cancellationToken);
        }
        var contentAssociations = new AssociatedFileSet();
        foreach (AssociatedFileOwner owner in associatedFileOwners) contentAssociations.UnionWith(owner.Files.Select(static entry => entry.Identity));

        if (names == null) { names = new PdfDictionary(); catalog.Items["Names"] = names; }
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        var nameEntries = new List<(byte[] KeyBytes, PdfReference FileSpec)>(attachments.Count);
        var fileSpecificationsByName = new Dictionary<string, PdfReference>(StringComparer.Ordinal);
        var fileSpecificationsByOriginalObject = new Dictionary<int, PdfReference>();
        var fileSpecificationsByOriginalEmbeddedFile = new Dictionary<int, PdfReference>();
        var associated = new PdfArray();
        foreach (var item in attachments.Select(static entry => (Entry: entry, KeyBytes: PdfTextString.Encode(entry.Attachment.FileName)))
            .OrderBy(static item => item.KeyBytes, PdfNameTreeKeyComparer.Instance)) {
            PdfAttachmentEditEntry entry = item.Entry;
            PdfEmbeddedFile attachment = entry.Attachment;
            byte[] data = attachment.DataSnapshot;
            int streamNumber = nextObjectNumber++;
            int fileSpecNumber = nextObjectNumber++;
            PdfDictionary streamDictionary = BuildEmbeddedFileDictionary(attachment, data);
            objects[streamNumber] = new PdfIndirectObject(streamNumber, 0, new PdfStream(streamDictionary, data));
            PdfDictionary fileSpec = BuildFileSpec(attachment, streamNumber);
            objects[fileSpecNumber] = new PdfIndirectObject(fileSpecNumber, 0, fileSpec);
            var reference = new PdfReference(fileSpecNumber, 0);
            nameEntries.Add((item.KeyBytes, reference));
            fileSpecificationsByName[attachment.FileName] = reference;
            if (entry.SourceIdentity.FileSpecObjectNumber > 0) {
                fileSpecificationsByOriginalObject[entry.SourceIdentity.FileSpecObjectNumber] = reference;
            }
            if (entry.SourceIdentity.EmbeddedFileObjectNumber > 0) {
                fileSpecificationsByOriginalEmbeddedFile[entry.SourceIdentity.EmbeddedFileObjectNumber] = reference;
            }
            // A file associated only with content stays a content association instead of becoming document-level.
            bool contentOnly = contentAssociations.Contains(entry.SourceIdentity) && !catalogAssociations.Contains(entry.SourceIdentity);
            if (attachment.Relationship != PdfAssociatedFileRelationship.Unspecified && !contentOnly) associated.Items.Add(reference);
        }
        var nameArray = new PdfArray();
        foreach ((byte[] keyBytes, PdfReference reference) in nameEntries) { nameArray.Items.Add(new PdfStringObj(keyBytes)); nameArray.Items.Add(reference); }
        var embeddedFiles = new PdfDictionary(); embeddedFiles.Items["Names"] = nameArray; names.Items["EmbeddedFiles"] = embeddedFiles;
        if (associated.Items.Count > 0) catalog.Items["AF"] = associated;
        ReconnectAssociatedFileOwners(associatedFileOwners, fileSpecificationsByOriginalObject, fileSpecificationsByOriginalEmbeddedFile);
        return ReconnectFileAttachmentAnnotations(
            objects,
            catalog,
            fileAttachmentAnnotations,
            retainedOriginalNames,
            fileSpecificationsByName,
            fileSpecificationsByOriginalObject,
            fileSpecificationsByOriginalEmbeddedFile,
            cancellationToken);
    }

    private static List<AssociatedFileOwner> CaptureAssociatedFileOwners(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        out AssociatedFileSet catalogAssociations,
        CancellationToken cancellationToken) {
        var owners = new List<AssociatedFileOwner>();
        catalogAssociations = new AssociatedFileSet();
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            CaptureAssociatedFileOwners(objects, indirect.Value, catalog, owners, catalogAssociations, visited);
        }
        return owners;
    }

    private static void CaptureAssociatedFileOwners(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfDictionary catalog,
        List<AssociatedFileOwner> owners,
        AssociatedFileSet catalogAssociations,
        ISet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            if (dictionary.Items.TryGetValue("AF", out PdfObject? associatedFiles) &&
                PdfObjectLookup.Resolve(objects, associatedFiles) is PdfArray files) {
                var entries = new List<AssociatedFileEntry>(files.Items.Count);
                foreach (PdfObject file in files.Items) entries.Add(new AssociatedFileEntry(file, ResolveFileSpecificationIdentity(objects, file)));
                if (ReferenceEquals(dictionary, catalog)) catalogAssociations.UnionWith(entries.Select(static entry => entry.Identity));
                else if (entries.Count > 0) owners.Add(new AssociatedFileOwner(dictionary, entries));
            }
            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) CaptureAssociatedFileOwners(objects, child, catalog, owners, catalogAssociations, visited);
            }
            return;
        }
        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) CaptureAssociatedFileOwners(objects, child, catalog, owners, catalogAssociations, visited);
            }
        }
    }

    private static PdfAttachmentSourceIdentity ResolveFileSpecificationIdentity(Dictionary<int, PdfIndirectObject> objects, PdfObject fileSpecification) {
        int fileSpecObjectNumber = fileSpecification is PdfReference reference ? reference.ObjectNumber : 0;
        PdfDictionary? embeddedFiles = ResolveDictionary(objects, ResolveDictionary(objects, fileSpecification) is PdfDictionary specification &&
            specification.Items.TryGetValue("EF", out PdfObject? embeddedFilesObject) ? embeddedFilesObject : null);
        PdfObject? embeddedFile = embeddedFiles == null ? null
            : embeddedFiles.Items.TryGetValue("UF", out PdfObject? unicodeFile) ? unicodeFile
            : embeddedFiles.Items.TryGetValue("F", out PdfObject? regularFile) ? regularFile : null;
        return new PdfAttachmentSourceIdentity(fileSpecObjectNumber, embeddedFile is PdfReference stream ? stream.ObjectNumber : 0);
    }

    private static void ReconnectAssociatedFileOwners(
        IReadOnlyList<AssociatedFileOwner> owners,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalObject,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalEmbeddedFile) {
        foreach (AssociatedFileOwner owner in owners) {
            var files = new PdfArray();
            var added = new HashSet<int>();
            foreach (AssociatedFileEntry entry in owner.Files) {
                PdfAttachmentSourceIdentity identity = entry.Identity;
                if (identity.EmbeddedFileObjectNumber <= 0) {
                    // External file references are not attachments; this edit leaves them untouched.
                    files.Items.Add(entry.Original);
                    continue;
                }
                PdfReference? replacement = null;
                if (identity.FileSpecObjectNumber > 0) fileSpecificationsByOriginalObject.TryGetValue(identity.FileSpecObjectNumber, out replacement);
                if (replacement == null && identity.EmbeddedFileObjectNumber > 0)
                    fileSpecificationsByOriginalEmbeddedFile.TryGetValue(identity.EmbeddedFileObjectNumber, out replacement);
                // Removed files leave no dangling association; retained and renamed files keep their owner.
                if (replacement != null && added.Add(replacement.ObjectNumber)) files.Items.Add(replacement);
            }
            if (files.Items.Count > 0) owner.Dictionary.Items["AF"] = files;
        }
    }

    private static void RemoveExistingEmbeddedFileGraph(Dictionary<int, PdfIndirectObject> objects, CancellationToken cancellationToken) {
        var removedObjectNumbers = new HashSet<int>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
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
        PdfDictionary catalog,
        IReadOnlyList<FileAttachmentAnnotation> annotations,
        IReadOnlyDictionary<string, string> retainedOriginalNames,
        Dictionary<string, PdfReference> fileSpecificationsByName,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalObject,
        Dictionary<int, PdfReference> fileSpecificationsByOriginalEmbeddedFile,
        CancellationToken cancellationToken) {
        var removedObjectNumbers = new HashSet<int>();
        var removedDirectAnnotations = new HashSet<PdfDictionary>();
        foreach (FileAttachmentAnnotation annotation in annotations) {
            PdfReference? replacement = null;
            if (annotation.FileSpecObjectNumber > 0) {
                fileSpecificationsByOriginalObject.TryGetValue(
                    annotation.FileSpecObjectNumber, out replacement);
            }
            if (replacement == null && annotation.EmbeddedFileObjectNumber > 0) {
                fileSpecificationsByOriginalEmbeddedFile.TryGetValue(
                    annotation.EmbeddedFileObjectNumber, out replacement);
            }
            if (replacement == null && annotation.FileSpecObjectNumber <= 0 &&
                annotation.EmbeddedFileObjectNumber <= 0 && annotation.FileName != null &&
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
        // Tagged annotations are structure content: drop their object references and parent-tree keys too.
        if (removedObjectNumbers.Count > 0 && catalog.Items.ContainsKey("StructTreeRoot"))
            PdfTaggedPageOrganizer.PruneStructure(objects, catalog, new HashSet<int>(), new HashSet<int>(removedObjectNumbers), cancellationToken);
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

    private sealed class AssociatedFileOwner {
        internal AssociatedFileOwner(PdfDictionary dictionary, IReadOnlyList<AssociatedFileEntry> files) {
            Dictionary = dictionary;
            Files = files;
        }

        internal PdfDictionary Dictionary { get; }
        internal IReadOnlyList<AssociatedFileEntry> Files { get; }
    }

    private sealed class AssociatedFileEntry {
        internal AssociatedFileEntry(PdfObject original, PdfAttachmentSourceIdentity identity) {
            Original = original;
            Identity = identity;
        }

        internal PdfObject Original { get; }
        internal PdfAttachmentSourceIdentity Identity { get; }
    }

    /// <summary>Matches a file by either its file-specification or embedded-stream object.</summary>
    private sealed class AssociatedFileSet {
        private readonly HashSet<int> _fileSpecifications = new();
        private readonly HashSet<int> _embeddedFiles = new();

        internal void UnionWith(IEnumerable<PdfAttachmentSourceIdentity> identities) {
            foreach (PdfAttachmentSourceIdentity identity in identities) {
                if (identity.FileSpecObjectNumber > 0) _fileSpecifications.Add(identity.FileSpecObjectNumber);
                if (identity.EmbeddedFileObjectNumber > 0) _embeddedFiles.Add(identity.EmbeddedFileObjectNumber);
            }
        }

        internal bool Contains(PdfAttachmentSourceIdentity identity) =>
            identity.FileSpecObjectNumber > 0 && _fileSpecifications.Contains(identity.FileSpecObjectNumber) ||
            identity.EmbeddedFileObjectNumber > 0 && _embeddedFiles.Contains(identity.EmbeddedFileObjectNumber);
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

    private static void ValidateAttachmentGraph(byte[] pdf, int expectedCount, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        int embeddedStreams = 0;
        int fileSpecifications = 0;
        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
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

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfAttachmentValidation> Validate(byte[] pdf, PdfEmbeddedFile[] target, PdfLoadOptions readOptions, CancellationToken cancellationToken) {
        IReadOnlyList<PdfExtractedAttachment> actual = PdfReadDocument.Open(pdf, readOptions, cancellationToken).ExtractAttachments(cancellationToken);
        var unmatched = actual.ToList();
        var result = new List<PdfAttachmentValidation>(target.Length);
        foreach (PdfEmbeddedFile expected in target) {
            cancellationToken.ThrowIfCancellationRequested();
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

    #pragma warning disable CA5351, CA1850 // PDF embedded-file /CheckSum is an MD5 file identifier mandated by the format; use Create for net472.
    private static byte[] ComputeChecksum(byte[] data) { using (MD5 md5 = MD5.Create()) return md5.ComputeHash(data); }
    #pragma warning restore CA5351, CA1850
    private static string ToHex(byte[] bytes) { var builder = new System.Text.StringBuilder(bytes.Length * 2); for (int i = 0; i < bytes.Length; i++) builder.Append(bytes[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture)); return builder.ToString(); }
}
