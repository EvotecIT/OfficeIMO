using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

/// <summary>Inspects and selectively removes the standards-defined PDF C2PA associated-file carrier.</summary>
public static class PdfProvenance {
    private const string C2paMimeType = "application/c2pa";

    /// <summary>Inspects a bounded PDF for C2PA Manifest Store associated files.</summary>
    public static OfficeProvenanceReport Inspect(
        byte[] pdf,
        OfficeProvenanceOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        if (pdf.LongLength > options.MaxAssetBytes) throw new InvalidDataException("The PDF exceeds the configured asset limit.");

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            Math.Min(options.MaxExpandedContainerBytes, MultiplySaturating(options.MaxManifestBytes, options.MaxCarriers)),
            options.MaxManifestBytes,
            options.MaxCarriers);
        PdfC2paAssociationProfile associations = CollectAssociationProfile(document);
        var evidence = new List<OfficeProvenanceEvidence>();
        foreach (PdfExtractedAttachment attachment in attachments) {
            if (!IsCandidate(attachment)) continue;
            byte[] manifest = attachment.Bytes;
            if (manifest.LongLength > options.MaxManifestBytes) throw new InvalidDataException("A PDF provenance manifest exceeds the configured manifest limit.");
            bool valid = attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest &&
                string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase) &&
                attachment.FileSpecObjectNumber > 0 &&
                IsFileSpecificationObject(document.Objects, attachment.FileSpecObjectNumber) &&
                associations.IsValid(attachment.FileSpecObjectNumber) &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
            if (evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
            evidence.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"PDF/Filespec[{attachment.FileSpecObjectNumber}]/{attachment.FileName}",
                valid,
                manifest.LongLength));
        }
        return new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Pdf, evidence.AsReadOnly());
    }

    /// <summary>Inspects a bounded PDF file.</summary>
    public static OfficeProvenanceReport InspectFile(
        string filePath,
        OfficeProvenanceOptions? options = null,
        PdfReadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        byte[] pdf = ReadBounded(filePath, options.MaxAssetBytes);
        return Inspect(pdf, options, readOptions);
    }

    /// <summary>Removes selected structurally valid C2PA associated files through the proven PDF attachment rewrite.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        byte[] pdf,
        OfficeProvenanceRemovalOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceReport before = Inspect(pdf, options.Limits, readOptions);
        if (!options.RemoveC2paManifests || before.Evidence.Count == 0) {
            return new OfficeProvenanceRemovalResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeProvenanceChange>(), false);
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            Math.Min(options.Limits.MaxExpandedContainerBytes, MultiplySaturating(options.Limits.MaxManifestBytes, options.Limits.MaxCarriers)),
            options.Limits.MaxManifestBytes,
            options.Limits.MaxCarriers);
        var removeFileSpecifications = new HashSet<int>();
        var changes = new List<OfficeProvenanceChange>();
        int evidenceIndex = 0;
        for (int index = 0; index < attachments.Count; index++) {
            PdfExtractedAttachment attachment = attachments[index];
            if (!IsCandidate(attachment)) continue;
            OfficeProvenanceEvidence evidence = before.Evidence[evidenceIndex++];
            if (!evidence.IsStructurallyValid && options.RequireStructurallyValidCarrier) continue;
            if (attachment.FileSpecObjectNumber <= 0) {
                throw new InvalidDataException("A direct PDF provenance filespec cannot be removed without risking unrelated associations.");
            }
            removeFileSpecifications.Add(attachment.FileSpecObjectNumber);
            changes.Add(new OfficeProvenanceChange(
                OfficeProvenanceCarrierKind.C2paManifest,
                evidence.Location,
                removedBytes: 0));
        }
        if (removeFileSpecifications.Count == 0) {
            return new OfficeProvenanceRemovalResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeProvenanceChange>(), false);
        }

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);
        if (security.HasSignatures) {
            string detail = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                ? "OfficeIMO.Pdf does not silently delete PDF signature revisions or fields; remove the signature through an explicit PDF signature workflow first."
                : "Removing provenance would invalidate PDF signatures.";
            throw new InvalidOperationException(detail);
        }

        byte[] output = PdfProvenanceGraphEditor.RemoveFileSpecifications(
            pdf,
            removeFileSpecifications,
            readOptions,
            options.Limits.MaxExpandedContainerBytes);
        PdfReadOptions outputReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, output.LongLength);
        OfficeProvenanceReport after = Inspect(output, options.Limits, outputReadOptions);
        return new OfficeProvenanceRemovalResult(output, before, after, changes.AsReadOnly(), true);
    }

    /// <summary>Removes selected provenance and atomically writes the resulting PDF.</summary>
    public static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null,
        PdfReadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        byte[] pdf = ReadBounded(inputPath, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(pdf, options, readOptions);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    private static bool IsCandidate(PdfExtractedAttachment attachment) =>
        attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest ||
        string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase);

    private static bool IsCandidate(PdfAttachmentInfo attachment) =>
        attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest ||
        string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase);

    private static PdfC2paAssociationProfile CollectAssociationProfile(PdfReadDocument document) {
        var documentLevel = new HashSet<int>();
        var objectLevel = new HashSet<int>();
        var secondaryDocumentReferences = new HashSet<int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(document.Objects, document.TrailerRaw);
        if (catalog == null) return new PdfC2paAssociationProfile(documentLevel, objectLevel, secondaryDocumentReferences);
        AddReferencesFromArray(document.Objects, catalog.Items.TryGetValue("AF", out PdfObject? catalogAf) ? catalogAf : null, documentLevel);
        CollectEmbeddedFilesNameTreeReferences(document.Objects, catalog, secondaryDocumentReferences);
        PdfIndirectObject catalogObject = document.Objects.Values.First(item => ReferenceEquals(item.Value, catalog));
        HashSet<int> reachableObjectNumbers = CollectReachableObjectNumbers(
            document.Objects,
            new PdfReference(catalogObject.ObjectNumber, catalogObject.Generation));
        HashSet<PdfObject> structuralAssociationSites = CollectEmbeddedFileStructuralDictionaries(
            document.Objects, reachableObjectNumbers);
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in document.Objects.Values.Where(item => reachableObjectNumbers.Contains(item.ObjectNumber))) {
            CollectObjectAssociations(document.Objects, item.Value, catalog, objectLevel, visited, structuralAssociationSites);
        }
        CollectPageAnnotationReferences(
            document.Objects,
            catalog,
            secondaryDocumentReferences,
            document.ReadOptions.Limits.MaxAnnotationsPerPage);
        return new PdfC2paAssociationProfile(documentLevel, objectLevel, secondaryDocumentReferences);
    }

    private static HashSet<int> CollectReachableObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReference root) {
        var result = new HashSet<int>();
        var visitedDirectObjects = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(root);
        while (pending.Count > 0) {
            PdfObject value = pending.Pop();
            if (value is PdfReference reference) {
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    !result.Add(reference.ObjectNumber)) continue;
                pending.Push(indirect.Value);
                continue;
            }
            if (!visitedDirectObjects.Add(value)) continue;
            PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
            if (dictionary != null) {
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            } else if (value is PdfArray array) {
                foreach (PdfObject child in array.Items) pending.Push(child);
            }
        }
        return result;
    }

    private static long MultiplySaturating(long value, int multiplier) =>
        value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;

    private static void CollectObjectAssociations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfDictionary catalog,
        HashSet<int> objectLevel,
        HashSet<PdfObject> visited,
        HashSet<PdfObject> structuralAssociationSites) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            if (!ReferenceEquals(dictionary, catalog) && IsInformationResource(value, dictionary, structuralAssociationSites)) {
                AddReferencesFromArray(objects, dictionary.Items.TryGetValue("AF", out PdfObject? associated) ? associated : null, objectLevel);
            }
            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) CollectObjectAssociations(
                    objects, child, catalog, objectLevel, visited, structuralAssociationSites);
            }
            return;
        }
        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) CollectObjectAssociations(
                    objects, child, catalog, objectLevel, visited, structuralAssociationSites);
            }
        }
    }

    private static bool IsInformationResource(
        PdfObject owner,
        PdfDictionary dictionary,
        HashSet<PdfObject> structuralAssociationSites) {
        if (structuralAssociationSites.Contains(dictionary)) return false;
        string? type = dictionary.Get<PdfName>("Type")?.Name;
        string? subtype = dictionary.Get<PdfName>("Subtype")?.Name;
        if (type is "Catalog" or "Pages" or "Page" or "Annot" or "Filespec" or "EmbeddedFile" or "XRef" or "ObjStm") return false;
        if (subtype is "FileAttachment" or "Popup" ||
            subtype != null && dictionary.Items.ContainsKey("Rect")) return false;
        if (dictionary.Items.ContainsKey("EF") &&
            (dictionary.Items.ContainsKey("F") || dictionary.Items.ContainsKey("UF"))) return false;
        if (owner is PdfStream) return true;
        return string.Equals(type, "StructElem", StringComparison.Ordinal) || type == null;
    }

    private static HashSet<PdfObject> CollectEmbeddedFileStructuralDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> reachableObjectNumbers) {
        var result = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values.Where(item => reachableObjectNumbers.Contains(item.ObjectNumber))) {
            PdfDictionary? dictionary = item.Value is PdfStream stream ? stream.Dictionary : item.Value as PdfDictionary;
            if (dictionary?.Get<PdfName>("Type")?.Name != "EmbeddedFile") continue;
            PdfObject? parameters = PdfObjectLookup.Resolve(
                objects,
                dictionary.Items.TryGetValue("Params", out PdfObject? value) ? value : null);
            if (parameters is PdfDictionary) result.Add(parameters);
        }
        return result;
    }

    private static void CollectPageAnnotationReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        HashSet<int> result,
        int maximumAnnotationsPerPage) {
        if (!catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return;
        var pending = new Stack<PdfObject>();
        var visited = new HashSet<PdfObject>();
        pending.Push(pages);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary dictionary || !visited.Add(resolved)) continue;
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (type == "Pages") {
                if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is PdfArray kids) {
                    foreach (PdfObject child in kids.Items) pending.Push(child);
                }
                continue;
            }
            if (type != "Page" || PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray annotations) continue;
            if (annotations.Items.Count > maximumAnnotationsPerPage) {
                throw new InvalidDataException("PDF page annotations exceed the configured per-page limit.");
            }
            foreach (PdfObject annotationValue in annotations.Items) {
                if (PdfObjectLookup.Resolve(objects, annotationValue) is not PdfDictionary annotation ||
                    !string.Equals(annotation.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal) ||
                    !annotation.Items.TryGetValue("FS", out PdfObject? fileSpecification) || fileSpecification is not PdfReference reference ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    !IsFileSpecificationValue(indirect.Value)) continue;
                result.Add(reference.ObjectNumber);
            }
        }
    }

    private static bool IsFileSpecificationObject(Dictionary<int, PdfIndirectObject> objects, int objectNumber) =>
        objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) && IsFileSpecificationValue(indirect.Value);

    private static bool IsFileSpecificationValue(PdfObject value) {
        if (value is not PdfDictionary dictionary) return false;
        string? type = dictionary.Get<PdfName>("Type")?.Name;
        return type == null || string.Equals(type, "Filespec", StringComparison.Ordinal);
    }

    private static void AddReferencesFromArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<int> result) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfArray array) return;
        foreach (PdfObject item in array.Items) {
            if (item is PdfReference reference &&
                PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                IsFileSpecificationValue(indirect.Value)) result.Add(reference.ObjectNumber);
        }
    }

    private static void CollectEmbeddedFilesNameTreeReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        HashSet<int> result) {
        if (PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is not PdfDictionary names ||
            !names.Items.TryGetValue("EmbeddedFiles", out PdfObject? embeddedFiles)) return;
        var visited = new HashSet<PdfObject>();
        CollectNameTreeReferences(objects, embeddedFiles, result, visited);
    }

    private static void CollectNameTreeReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> result,
        HashSet<PdfObject> visited) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved == null || !visited.Add(resolved) || resolved is not PdfDictionary dictionary) return;
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is PdfArray names) {
            for (int index = 1; index < names.Items.Count; index += 2) {
                if (PdfObjectLookup.Resolve(objects, names.Items[index - 1]) is PdfStringObj &&
                    names.Items[index] is PdfReference reference &&
                    PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                    IsFileSpecificationValue(indirect.Value)) result.Add(reference.ObjectNumber);
            }
        }
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) return;
        foreach (PdfObject child in kids.Items) CollectNameTreeReferences(objects, child, result, visited);
    }

    private sealed class PdfC2paAssociationProfile {
        private readonly HashSet<int> _documentLevel;
        private readonly HashSet<int> _objectLevel;
        private readonly HashSet<int> _secondaryDocumentReferences;

        internal PdfC2paAssociationProfile(HashSet<int> documentLevel, HashSet<int> objectLevel, HashSet<int> secondaryDocumentReferences) {
            _documentLevel = documentLevel;
            _objectLevel = objectLevel;
            _secondaryDocumentReferences = secondaryDocumentReferences;
        }

        internal bool IsValid(int fileSpecObjectNumber) => fileSpecObjectNumber > 0 &&
            (_objectLevel.Contains(fileSpecObjectNumber) ||
             _documentLevel.Contains(fileSpecObjectNumber) && _secondaryDocumentReferences.Contains(fileSpecObjectNumber));
    }

    private static byte[] ReadBounded(string filePath, long maximumBytes) {
        using var stream = File.OpenRead(Path.GetFullPath(filePath));
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes);
    }
}
