namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    private static PdfMergeResult ApplyMergePolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeOptions? options) {
        PdfMergePolicy policy = options?.Policy ?? new PdfMergePolicy();
        Guard.NotNull(policy, nameof(policy));

        var inventories = sources.Select((source, index) => new PdfMergeSourceInventory(
            index,
            source.Document.Pages.Count,
            CountOutlines(source.Document.Outlines),
            source.Document.NamedDestinations.Count,
            source.Document.PageLabels.Count,
            source.Document.FormFields.Count,
            PdfAttachmentExtractor.ExtractAttachments(source.Document).Count)).ToArray();
        var decisions = new List<PdfMergeDecision>();
        if (options?.FlattenVisualAnnotations == true) decisions.Add(new PdfMergeDecision("SourceAnnotations", PdfMergeStructureMode.Combine, "Flattened supported visual annotations before import; links, forms, and unsupported shapes remained live."));
        if (options?.ResizePages != null) decisions.Add(new PdfMergeDecision("PageSizeNormalization", PdfMergeStructureMode.Combine, "Normalized every source page through the requested resize mode before import."));

        // Form imports rely on the source-to-output object map produced by the initial merge.
        // Apply them before any policy editor can rewrite and renumber the object graph.
        merged = ApplyFormPolicy(merged, sources, primarySourceIndex, policy.Forms, policy.FormFieldCollisions, decisions);
        merged = ApplyCatalogStatePolicy(merged, sources, primarySourceIndex, policy.CatalogState, decisions);
        merged = ApplyMetadataPolicy(merged, sources, primarySourceIndex, policy.Metadata, decisions);
        merged = ApplyNamedDestinationPolicy(merged, sources, primarySourceIndex, policy.NamedDestinations, policy.NamedDestinationCollisions, decisions);
        merged = ApplyPageLabelPolicy(merged, sources, primarySourceIndex, policy.PageLabels, decisions);
        merged = ApplyOutlinePolicy(merged, sources, primarySourceIndex, policy.Outlines, decisions);
        merged = ApplyAttachmentPolicy(merged, sources, primarySourceIndex, policy.Attachments, policy.AttachmentCollisions, decisions);
        merged = ApplyViewerPolicy(merged, sources, primarySourceIndex, policy.ViewerPreferences, decisions);

        PdfReadDocument readback = PdfReadDocument.Open(merged);
        int expectedPageCount = sources.Sum(static source => source.PageObjectNumbers.Length);
        if (readback.Pages.Count != expectedPageCount) {
            throw new InvalidOperationException("PDF merge post-save validation failed: output page count did not match the imported page count.");
        }

        return new PdfMergeResult(merged, new PdfMergeReport(inventories, decisions.AsReadOnly(), readback.Pages.Count));
    }

    private static byte[] ApplyMetadataPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        List<PdfMergeDecision> decisions) {
        PdfMetadata primary = sources[primarySourceIndex].Metadata;
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                decisions.Add(new PdfMergeDecision("Metadata", mode, "Kept primary metadata."));
                return merged;
            case PdfMergeStructureMode.Drop:
                decisions.Add(new PdfMergeDecision("Metadata", mode, "Removed document information and synchronized XMP metadata."));
                return PdfMetadataEditor.ReplaceMetadata(merged, new PdfMetadata());
            case PdfMergeStructureMode.RejectIncoming:
                int conflicts = sources.Where((source, index) => index != primarySourceIndex && HasMetadata(source.Metadata)).Count();
                if (conflicts > 0) throw new InvalidOperationException("PDF merge policy rejected incoming metadata from " + conflicts + " source(s).");
                decisions.Add(new PdfMergeDecision("Metadata", mode, "No incoming metadata was present."));
                return merged;
            case PdfMergeStructureMode.Combine:
                PdfMetadata combined = CombineMetadata(primary, sources, primarySourceIndex);
                int supplied = sources.Where((source, index) => index != primarySourceIndex && HasMetadata(source.Metadata)).Count();
                decisions.Add(new PdfMergeDecision("Metadata", mode, "Filled empty primary metadata values from later sources.", supplied));
                return PdfMetadataEditor.ReplaceMetadata(merged, combined);
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static byte[] ApplyOutlinePolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        List<PdfMergeDecision> decisions) {
        int incomingCount = sources.Where((source, index) => index != primarySourceIndex).Sum(source => CountOutlines(source.Document.Outlines));
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                decisions.Add(new PdfMergeDecision("Outlines", mode, "Kept the primary outline tree.", droppedCount: incomingCount));
                return merged;
            case PdfMergeStructureMode.RejectIncoming:
                if (incomingCount > 0) throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming outline node(s).");
                decisions.Add(new PdfMergeDecision("Outlines", mode, "No incoming outline nodes were present."));
                return merged;
            case PdfMergeStructureMode.Drop:
                merged = PdfBookmarkEditor.Edit(merged, session => RemoveAllBookmarks(session)).ToBytes();
                decisions.Add(new PdfMergeDecision("Outlines", mode, "Removed all outline trees."));
                return merged;
            case PdfMergeStructureMode.Combine:
                var roots = BuildMergedOutlineRoots(sources);
                merged = PdfBookmarkEditor.EditAllowingBrokenSourceDestinations(merged, session => {
                    RemoveAllBookmarks(session);
                    foreach (MergedOutlineNode root in roots) AddBookmark(session, root, null);
                }).ToBytes();
                decisions.Add(new PdfMergeDecision("Outlines", mode, "Appended source outline roots in source order and retargeted them to merged pages.", incomingCount));
                return merged;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static byte[] ApplyAttachmentPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<PdfMergeDecision> decisions) {
        IReadOnlyList<PdfExtractedAttachment>[] sourceAttachments = sources.Select(source => PdfAttachmentExtractor.ExtractAttachments(source.Document)).ToArray();
        int incomingCount = sourceAttachments.Where((_, index) => index != primarySourceIndex).Sum(static items => items.Count);
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                if (incomingCount > 0) merged = ReplaceAttachments(merged, ConvertAttachments(sourceAttachments[primarySourceIndex]));
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Kept primary attachments.", droppedCount: incomingCount));
                return merged;
            case PdfMergeStructureMode.RejectIncoming:
                if (incomingCount > 0) throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming attachment(s).");
                decisions.Add(new PdfMergeDecision("Attachments", mode, "No incoming attachments were present."));
                return merged;
            case PdfMergeStructureMode.Drop:
                if (sourceAttachments.Sum(static items => items.Count) > 0) merged = ReplaceAttachments(merged, Array.Empty<PdfEmbeddedFile>());
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Removed embedded and associated files."));
                return merged;
            case PdfMergeStructureMode.Combine:
                var renamed = new List<string>();
                int dropped = 0;
                IReadOnlyList<PdfEmbeddedFile> combined = CombineAttachments(sourceAttachments, collisionMode, renamed, ref dropped);
                merged = ReplaceAttachments(merged, combined);
                decisions.Add(new PdfMergeDecision("Attachments", mode, "Combined attachments and validated payload and metadata readback.", incomingCount - dropped, dropped, renamed.AsReadOnly()));
                return merged;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfEmbeddedFile> CombineAttachments(
        IReadOnlyList<PdfExtractedAttachment>[] sources,
        PdfMergeCollisionMode collisionMode,
        List<string> renamed,
        ref int dropped) {
        var result = new List<PdfEmbeddedFile>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        for (int sourceIndex = 0; sourceIndex < sources.Length; sourceIndex++) {
            foreach (PdfExtractedAttachment attachment in sources[sourceIndex]) {
                string name = attachment.FileName;
                if (!names.Add(name)) {
                    if (collisionMode == PdfMergeCollisionMode.Reject) throw new InvalidOperationException("PDF attachment name collision: " + name);
                    if (collisionMode == PdfMergeCollisionMode.KeepFirst) { dropped++; continue; }
                    string renamedName = GetUniqueIncomingName(name, sourceIndex, names);
                    renamed.Add("source " + sourceIndex + ": " + name + " -> " + renamedName);
                    name = renamedName;
                    names.Add(name);
                }
                result.Add(new PdfEmbeddedFile(name, attachment.Bytes, attachment.MimeType, attachment.Relationship, attachment.Description, attachment.CreationDate, attachment.ModificationDate));
            }
        }
        return result.AsReadOnly();
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

    private static byte[] ReplaceAttachments(byte[] merged, IReadOnlyList<PdfEmbeddedFile> attachments) {
        return PdfAttachmentEditor.Edit(merged, session => {
            foreach (PdfEmbeddedFile attachment in session.Attachments.ToArray()) session.Remove(attachment.FileName);
            foreach (PdfEmbeddedFile attachment in attachments) session.Add(attachment);
        }).ToBytes();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfEmbeddedFile> ConvertAttachments(IReadOnlyList<PdfExtractedAttachment> attachments) {
        return attachments.Select(static attachment => new PdfEmbeddedFile(
            attachment.FileName,
            attachment.Bytes,
            attachment.MimeType,
            attachment.Relationship,
            attachment.Description,
            attachment.CreationDate,
            attachment.ModificationDate)).ToList().AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<MergedOutlineNode> BuildMergedOutlineRoots(IReadOnlyList<ImportedSource> sources) {
        var result = new List<MergedOutlineNode>();
        int pageOffset = 0;
        foreach (ImportedSource source in sources) {
            foreach (PdfOutlineItem outline in source.Document.Outlines) {
                ImportOutline(result, outline, pageOffset);
            }
            pageOffset += source.PageObjectNumbers.Length;
        }
        return result.AsReadOnly();
    }

    private static void ImportOutline(List<MergedOutlineNode> target, PdfOutlineItem outline, int pageOffset) {
        var children = new List<MergedOutlineNode>();
        foreach (PdfOutlineItem child in outline.Children) {
            ImportOutline(children, child, pageOffset);
        }
        if (!outline.PageNumber.HasValue) {
            target.AddRange(children);
            return;
        }
        target.Add(new MergedOutlineNode(outline.Title, outline.PageNumber.Value + pageOffset, outline.DestinationTop, outline.IsExpanded, children.AsReadOnly()));
    }

    private static void AddBookmark(PdfBookmarkEditSession session, MergedOutlineNode node, string? parentId) {
        PdfBookmarkNode added = session.Add(node.Title, node.PageNumber, parentId, node.DestinationTop, node.IsExpanded);
        foreach (MergedOutlineNode child in node.Children) AddBookmark(session, child, added.Id);
    }

    private static void RemoveAllBookmarks(PdfBookmarkEditSession session) {
        foreach (PdfBookmarkNode root in session.Roots.ToArray()) session.Remove(root.Id);
    }

    private static PdfMetadata CombineMetadata(PdfMetadata primary, IReadOnlyList<ImportedSource> sources, int primarySourceIndex) {
        string? title = primary.Title; string? author = primary.Author; string? subject = primary.Subject; string? keywords = primary.Keywords;
        for (int i = 0; i < sources.Count; i++) {
            if (i == primarySourceIndex) continue;
            PdfMetadata incoming = sources[i].Metadata;
            if (string.IsNullOrEmpty(title)) title = incoming.Title;
            if (string.IsNullOrEmpty(author)) author = incoming.Author;
            if (string.IsNullOrEmpty(subject)) subject = incoming.Subject;
            if (string.IsNullOrEmpty(keywords)) keywords = incoming.Keywords;
        }
        return new PdfMetadata { Title = title, Author = author, Subject = subject, Keywords = keywords };
    }

    private static bool HasMetadata(PdfMetadata metadata) =>
        !string.IsNullOrEmpty(metadata.Title) || !string.IsNullOrEmpty(metadata.Author) ||
        !string.IsNullOrEmpty(metadata.Subject) || !string.IsNullOrEmpty(metadata.Keywords);

    private static int CountOutlines(IReadOnlyList<PdfOutlineItem> outlines) {
        int count = 0;
        foreach (PdfOutlineItem outline in outlines) count += 1 + CountOutlines(outline.Children);
        return count;
    }

    private sealed class MergedOutlineNode {
        internal MergedOutlineNode(string title, int pageNumber, double? destinationTop, bool isExpanded, IReadOnlyList<MergedOutlineNode> children) {
            Title = title; PageNumber = pageNumber; DestinationTop = destinationTop; IsExpanded = isExpanded; Children = children;
        }
        internal string Title { get; }
        internal int PageNumber { get; }
        internal double? DestinationTop { get; }
        internal bool IsExpanded { get; }
        internal IReadOnlyList<MergedOutlineNode> Children { get; }
    }
}
