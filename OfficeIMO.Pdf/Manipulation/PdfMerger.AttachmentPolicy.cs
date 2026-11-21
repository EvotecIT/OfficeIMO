namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    /// <summary>
    /// Applies the attachment policy to the merged document through the shared attachment editor. Files that
    /// belong to page or structure content travel with that content; retained files keep their object identity,
    /// so structure-element associations and file attachment annotations stay linked after renames.
    /// </summary>
    private static byte[] ApplyAttachmentPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<PdfMergeDecision> decisions,
        PdfLoadOptions readOptions,
        int[]? outputSourceIndexes,
        CancellationToken cancellationToken) {
        int[] sourceAttachmentCounts = sources.Select(static source => source.Document.Attachments.Count).ToArray();
        int incomingCount = sourceAttachmentCounts.Where((_, index) => index != primarySourceIndex).Sum();
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                if (incomingCount > 0) {
                    merged = EditMergedAttachments(merged, sources, primarySourceIndex, readOptions, outputSourceIndexes, (session, owners) => {
                        IReadOnlyList<PdfAttachmentEditEntry> entries = session.SnapshotEntries();
                        session.Recompose(entries
                            .Select((entry, index) => (Entry: entry, Index: index))
                            .Where(item => owners[item.Index] == primarySourceIndex)
                            .Select(item => ((int?)item.Index, item.Entry.Attachment))
                            .ToArray());
                    }, cancellationToken);
                }
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Kept primary attachments.", droppedCount: incomingCount));
                return merged;
            case PdfMergeStructureMode.RejectIncoming:
                if (incomingCount > 0) throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming attachment(s).");
                decisions.Add(new PdfMergeDecision("Attachments", mode, "No incoming attachments were present."));
                return merged;
            case PdfMergeStructureMode.Drop:
                if (sourceAttachmentCounts.Sum() > 0) {
                    merged = EditMergedAttachments(merged, sources, primarySourceIndex, readOptions, outputSourceIndexes,
                        (session, _) => session.Recompose(Array.Empty<(int?, PdfEmbeddedFile)>()), cancellationToken);
                }
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Removed embedded and associated files."));
                return merged;
            case PdfMergeStructureMode.Combine:
                var renamed = new List<string>();
                int dropped = 0;
                IReadOnlyList<PdfExtractedAttachment>[] documentAttachments = sources
                    .Select((source, index) => index == primarySourceIndex
                        ? Array.Empty<PdfExtractedAttachment>()
                        : ExtractDocumentLevelAttachments(source, cancellationToken))
                    .ToArray();
                merged = EditMergedAttachments(merged, sources, primarySourceIndex, readOptions, outputSourceIndexes, (session, owners) =>
                    session.Recompose(CombineAttachments(session.SnapshotEntries(), owners, documentAttachments, collisionMode, renamed, ref dropped)),
                    cancellationToken);
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Combined attachments and validated payload and metadata readback.", incomingCount - dropped, dropped, renamed.AsReadOnly()));
                return merged;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static byte[] EditMergedAttachments(byte[] merged, IReadOnlyList<ImportedSource> sources, int primarySourceIndex,
        PdfLoadOptions readOptions, int[]? outputSourceIndexes, Action<PdfAttachmentEditSession, IReadOnlyList<int>> edit, CancellationToken cancellationToken) {
        PdfReadDocument read = PdfReadDocument.Open(merged, readOptions, cancellationToken);
        int[] pageOwners = outputSourceIndexes ?? sources
            .SelectMany((source, index) => Enumerable.Repeat(index, source.PageObjectNumbers.Length))
            .ToArray();
        if (pageOwners.Length != read.Pages.Count) throw new InvalidOperationException("PDF merge output page ownership does not match the merged page count.");
        var ownerByPage = new Dictionary<int, int>();
        for (int index = 0; index < read.Pages.Count; index++) ownerByPage[read.Pages[index].ObjectNumber] = pageOwners[index];
        Dictionary<int, int> contentOwners = MapContentAttachmentOwners(read, ownerByPage, cancellationToken);
        return PdfAttachmentEditor.Edit(merged, session => {
            // Files outside page and structure content belong to the merged catalog, which is the primary source's.
            int[] owners = session.SnapshotEntries()
                .Select(entry => contentOwners.TryGetValue(entry.SourceIdentity.EmbeddedFileObjectNumber, out int owner) ? owner : primarySourceIndex)
                .ToArray();
            edit(session, owners);
        }, readOptions, cancellationToken: cancellationToken).ToBytes();
    }

    /// <summary>Returns the source's catalog-level attachments; content-bound files arrive with the imported pages.</summary>
    private static PdfExtractedAttachment[] ExtractDocumentLevelAttachments(ImportedSource source, CancellationToken cancellationToken) {
        if (source.Document.Attachments.Count == 0) return Array.Empty<PdfExtractedAttachment>();
        var importedPages = new Dictionary<int, int>();
        foreach (int page in source.PageObjectNumbers) importedPages[page] = 0;
        Dictionary<int, int> contentFiles = MapContentAttachmentOwners(source.Document, importedPages, cancellationToken);
        return source.Document.ExtractAttachments(cancellationToken)
            .Where(attachment => !contentFiles.ContainsKey(attachment.EmbeddedFileObjectNumber))
            .ToArray();
    }

    private static (int?, PdfEmbeddedFile)[] CombineAttachments(
        IReadOnlyList<PdfAttachmentEditEntry> mergedEntries,
        IReadOnlyList<int> owners,
        IReadOnlyList<PdfExtractedAttachment>[] documentAttachments,
        PdfMergeCollisionMode collisionMode,
        List<string> renamed,
        ref int dropped) {
        var result = new List<(int?, PdfEmbeddedFile)>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        for (int sourceIndex = 0; sourceIndex < documentAttachments.Length; sourceIndex++) {
            var candidates = new List<(int? ExistingIndex, PdfEmbeddedFile Attachment)>();
            for (int index = 0; index < mergedEntries.Count; index++) {
                if (owners[index] == sourceIndex) candidates.Add((index, mergedEntries[index].Attachment));
            }
            foreach (PdfExtractedAttachment attachment in documentAttachments[sourceIndex]) {
                candidates.Add((null, new PdfEmbeddedFile(attachment.FileName, attachment.Bytes, attachment.MimeType, attachment.Relationship,
                    attachment.Description, attachment.CreationDate, attachment.ModificationDate)));
            }
            foreach ((int? existingIndex, PdfEmbeddedFile attachment) in candidates) {
                string name = attachment.FileName;
                if (!names.Add(name)) {
                    if (collisionMode == PdfMergeCollisionMode.Reject) throw new InvalidOperationException("PDF attachment name collision: " + name);
                    if (collisionMode == PdfMergeCollisionMode.KeepFirst) { dropped++; continue; }
                    string renamedName = GetUniqueIncomingName(name, sourceIndex, names);
                    renamed.Add("source " + sourceIndex + ": " + name + " -> " + renamedName);
                    names.Add(renamedName);
                    attachment.FileName = renamedName;
                }
                result.Add((existingIndex, attachment));
            }
        }
        return result.ToArray();
    }

    private static string GetUniqueIncomingName(string fileName, int sourceIndex, HashSet<string> names) {
        string extension = Path.GetExtension(fileName);
        string stem = Path.GetFileNameWithoutExtension(fileName);
        int sequence = 1;
        while (true) {
            string candidate = stem + ".source" + (sourceIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                (sequence == 1 ? string.Empty : "." + sequence.ToString(System.Globalization.CultureInfo.InvariantCulture)) + extension;
            if (!names.Contains(candidate)) return candidate;
            sequence++;
        }
    }

    /// <summary>
    /// Maps embedded-file streams reached from page content (page and annotation associations, file attachment
    /// annotations, form XObjects) and from structure elements to the owner of the page they belong to.
    /// </summary>
    private static Dictionary<int, int> MapContentAttachmentOwners(PdfReadDocument document,
        Dictionary<int, int> ownerByPage, CancellationToken cancellationToken) {
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        var owners = new Dictionary<int, int>();
        void Claim(PdfObject? fileSpecification, int owner) {
            if (PdfObjectLookup.Resolve(objects, fileSpecification) is not PdfDictionary specification ||
                PdfObjectLookup.Resolve(objects, specification.Items.TryGetValue("EF", out PdfObject? value) ? value : null) is not PdfDictionary files) return;
            PdfObject? stream = files.Items.TryGetValue("UF", out PdfObject? unicode) ? unicode : files.Items.TryGetValue("F", out PdfObject? file) ? file : null;
            if (stream is PdfReference reference && !owners.ContainsKey(reference.ObjectNumber)) owners.Add(reference.ObjectNumber, owner);
        }
        void ClaimAssociated(PdfDictionary? dictionary, int owner) {
            if (dictionary == null || PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("AF", out PdfObject? value) ? value : null) is not PdfArray files) return;
            foreach (PdfObject file in files.Items) Claim(file, owner);
        }
        foreach (PdfReadPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ownerByPage.TryGetValue(page.ObjectNumber, out int owner)) continue;
            PdfDictionary pageDictionary = page.PageDictionary;
            ClaimAssociated(pageDictionary, owner);
            if (PdfObjectLookup.Resolve(objects, pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotations) ? annotations : null) is PdfArray annotationArray) {
                foreach (PdfObject item in annotationArray.Items) {
                    if (PdfObjectLookup.Resolve(objects, item) is not PdfDictionary annotation) continue;
                    ClaimAssociated(annotation, owner);
                    if (annotation.Get<PdfName>("Subtype")?.Name == "FileAttachment") Claim(annotation.Items.TryGetValue("FS", out PdfObject? fs) ? fs : null, owner);
                }
            }
            if (PdfObjectLookup.Resolve(objects, pageDictionary.Items.TryGetValue("Resources", out PdfObject? resources) ? resources : null) is PdfDictionary resourceDictionary &&
                PdfObjectLookup.Resolve(objects, resourceDictionary.Items.TryGetValue("XObject", out PdfObject? xobjects) ? xobjects : null) is PdfDictionary xobjectDictionary) {
                foreach (PdfObject item in xobjectDictionary.Items.Values) {
                    if (PdfObjectLookup.Resolve(objects, item) is PdfStream xobject) ClaimAssociated(xobject.Dictionary, owner);
                }
            }
        }

        if (document.Security.RootObjectNumber is not int root || objects[root].Value is not PdfDictionary catalog ||
            PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? treeRoot) ? treeRoot : null) is not PdfDictionary structureRoot ||
            !structureRoot.Items.TryGetValue("K", out PdfObject? kids)) return owners;
        var visited = new HashSet<PdfDictionary>();
        // Returns the first page reached in a structure subtree. An element without its own /Pg uses an inherited
        // page, then the first page of its content, so container elements such as formulas resolve to their page.
        int? Visit(PdfObject value, int? inheritedPage, int depth) {
            cancellationToken.ThrowIfCancellationRequested();
            if (depth > 256) throw new InvalidOperationException("The tagged structure tree exceeds the supported nesting depth.");
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfArray array) {
                int? first = null;
                foreach (PdfObject item in array.Items) {
                    int? page = Visit(item, inheritedPage, depth + 1);
                    first ??= page;
                }
                return first;
            }
            if (resolved is PdfNumber) return inheritedPage;
            if (resolved is not PdfDictionary element || !visited.Add(element)) return null;
            int? ownPage = element.Items.TryGetValue("Pg", out PdfObject? pageValue) && pageValue is PdfReference pageReference ? pageReference.ObjectNumber : null;
            int? contentPage = element.Items.TryGetValue("K", out PdfObject? children) ? Visit(children, ownPage ?? inheritedPage, depth + 1) : null;
            int? ownerPage = ownPage ?? inheritedPage ?? contentPage;
            if (ownerPage is int pageNumber && ownerByPage.TryGetValue(pageNumber, out int owner)) ClaimAssociated(element, owner);
            return ownPage ?? contentPage;
        }
        Visit(kids, null, 0);
        return owners;
    }
}
