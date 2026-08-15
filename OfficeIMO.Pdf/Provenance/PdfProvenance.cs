using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

/// <summary>Inspects and selectively removes the standards-defined PDF C2PA associated-file carrier.</summary>
public static partial class PdfProvenance {
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

        long maximumManifestBytes = GetMaximumManifestBytes(options);
        PdfReadOptions effectiveReadOptions = CreateReadOptionsForInspection(options, readOptions);
        PdfReadDocument document = PdfReadDocument.Open(pdf, effectiveReadOptions);
        HashSet<int> pageTreeObjectNumbers = CollectPageTreeObjectNumbers(document, options.MaxContainerEntries);
        PdfC2paAssociationProfile associations = CollectAssociationProfile(
            document,
            pageTreeObjectNumbers,
            options.MaxContainerEntries,
            out HashSet<int> reachableObjectNumbers);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            maximumManifestBytes,
            options.MaxManifestBytes,
            options.MaxCarriers,
            options.MaxContainerEntries,
            requireSuccessfulDecoding: true,
            allowedObjectNumbers: reachableObjectNumbers);
        var evidence = new List<OfficeProvenanceEvidence>();
        foreach (PdfExtractedAttachment attachment in attachments) {
            if (!IsCandidate(attachment)) continue;
            byte[] manifest = attachment.Bytes;
            if (manifest.LongLength > options.MaxManifestBytes) throw new InvalidDataException("A PDF provenance manifest exceeds the configured manifest limit.");
            bool valid = attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest &&
                string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase) &&
                attachment.FileSpecObjectNumber > 0 &&
                HasEmbeddedFileStreamType(document.Objects, attachment) &&
                IsFileSpecificationObject(document.Objects, attachment.FileSpecObjectNumber, pageTreeObjectNumbers, associations.StructuralObjectNumbers) &&
                HasOnlySelectedEmbeddedFileVariants(document.Objects, attachment) &&
                associations.IsValid(attachment.FileSpecObjectNumber) &&
                OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
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
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        long maximumManifestBytes = Math.Min(
            options.Limits.MaxExpandedContainerBytes,
            MultiplySaturating(options.Limits.MaxManifestBytes, options.Limits.MaxCarriers));
        PdfReadOptions effectiveReadOptions = CreateReadOptions(
            options.Limits.MaxAssetBytes,
            options.Limits.MaxContainerEntries,
            options.Limits.MaxExpandedContainerBytes,
            options.Limits.MaxManifestBytes,
            maximumManifestBytes,
            readOptions);
        OfficeProvenanceReport before = Inspect(pdf, options.Limits, effectiveReadOptions);
        if (!options.RemoveC2paManifests || before.Evidence.Count == 0) {
            return new OfficeProvenanceRemovalResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeProvenanceChange>(), false);
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, effectiveReadOptions);
        HashSet<int> pageTreeObjectNumbers = CollectPageTreeObjectNumbers(document, options.Limits.MaxContainerEntries);
        _ = CollectAssociationProfile(
            document,
            pageTreeObjectNumbers,
            options.Limits.MaxContainerEntries,
            out HashSet<int> reachableObjectNumbers);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            maximumManifestBytes,
            options.Limits.MaxManifestBytes,
            options.Limits.MaxCarriers,
            options.Limits.MaxContainerEntries,
            requireSuccessfulDecoding: true,
            allowedObjectNumbers: reachableObjectNumbers);
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

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, effectiveReadOptions);
        if (security.HasEncryption) {
            throw new InvalidOperationException("Provenance removal does not remove or replace PDF encryption. Decrypt the document through an explicit PDF security workflow first.");
        }
        if (security.HasSignatures) {
            string detail = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                ? "OfficeIMO.Pdf does not silently delete PDF signature revisions or fields; remove the signature through an explicit PDF signature workflow first."
                : "Removing provenance would invalidate PDF signatures.";
            throw new InvalidOperationException(detail);
        }

        byte[] output = PdfProvenanceGraphEditor.RemoveFileSpecifications(
            pdf,
            removeFileSpecifications,
            effectiveReadOptions,
            options.Limits.MaxExpandedContainerBytes);
        PdfReadOptions outputReadOptions = PdfReadOptions.WithMinimumInputBytes(effectiveReadOptions, output.LongLength);
        OfficeProvenanceOptions outputLimits = CreateOutputInspectionOptions(options.Limits, output.LongLength);
        OfficeProvenanceReport after = Inspect(output, outputLimits, outputReadOptions);
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

    private static PdfC2paAssociationProfile CollectAssociationProfile(
        PdfReadDocument document,
        HashSet<int> pageTreeObjectNumbers,
        int maximumContainerEntries,
        out HashSet<int> reachableObjectNumbers) {
        var documentLevel = new HashSet<int>();
        var objectLevel = new HashSet<int>();
        var secondaryDocumentReferences = new HashSet<int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(document.Objects, document.TrailerRaw);
        if (catalog == null) {
            reachableObjectNumbers = new HashSet<int>();
            return new PdfC2paAssociationProfile(documentLevel, objectLevel, secondaryDocumentReferences, new HashSet<int>());
        }
        AddReferencesFromArray(document.Objects, catalog.Items.TryGetValue("AF", out PdfObject? catalogAf) ? catalogAf : null, documentLevel);
        CollectEmbeddedFilesNameTreeReferences(document.Objects, catalog, secondaryDocumentReferences, maximumContainerEntries);
        PdfIndirectObject catalogObject = document.Objects.Values.First(item => ReferenceEquals(item.Value, catalog));
        HashSet<int> reachableObjects = CollectReachableObjectNumbers(
            document.Objects,
            new PdfReference(catalogObject.ObjectNumber, catalogObject.Generation),
            maximumContainerEntries);
        reachableObjectNumbers = reachableObjects;
        HashSet<PdfObject> structuralAssociationSites = CollectStructuralAssociationSites(
            document.Objects,
            catalog,
            PdfSyntax.TryReadFirstReference(document.TrailerRaw, "Info"),
            document.Security.EncryptObjectNumber,
            reachableObjects,
            pageTreeObjectNumbers,
            new HashSet<int>(document.Pages.Select(static page => page.ObjectNumber)),
            maximumContainerEntries);
        var structuralObjectNumbers = new HashSet<int>(document.Objects.Values
            .Where(item => reachableObjects.Contains(item.ObjectNumber))
            .Where(item => {
                PdfDictionary? dictionary = item.Value is PdfStream stream ? stream.Dictionary : item.Value as PdfDictionary;
                return dictionary != null && structuralAssociationSites.Contains(dictionary);
            })
            .Select(item => item.ObjectNumber));
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in document.Objects.Values.Where(item => reachableObjects.Contains(item.ObjectNumber))) {
            CollectObjectAssociations(document.Objects, item.Value, catalog, objectLevel, visited, structuralAssociationSites);
        }
        CollectPageAnnotationReferences(
            document.Objects,
            catalog,
            secondaryDocumentReferences,
            structuralObjectNumbers,
            document.ReadOptions.Limits.MaxAnnotationsPerPage);
        return new PdfC2paAssociationProfile(documentLevel, objectLevel, secondaryDocumentReferences, structuralObjectNumbers);
    }

    private static HashSet<int> CollectReachableObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReference root,
        int maximumContainerEntries) {
        var result = new HashSet<int>();
        var visitedDirectObjects = new HashSet<PdfObject>();
        var indirectValues = new HashSet<PdfObject>(objects.Values.Select(static item => item.Value));
        int directStructuralEntries = 0;
        var pending = new Stack<PdfObject>();
        pending.Push(root);
        while (pending.Count > 0) {
            PdfObject value = pending.Pop();
            if (value is PdfReference reference) {
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    !result.Add(reference.ObjectNumber)) continue;
                if (result.Count > maximumContainerEntries - directStructuralEntries) {
                    throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
                }
                pending.Push(indirect.Value);
                continue;
            }
            if (!visitedDirectObjects.Add(value)) continue;
            PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
            if (dictionary != null) {
                if (!indirectValues.Contains(value) && ++directStructuralEntries > maximumContainerEntries - result.Count) {
                    throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
                }
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            } else if (value is PdfArray array) {
                if (!indirectValues.Contains(value) && ++directStructuralEntries > maximumContainerEntries - result.Count) {
                    throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
                }
                foreach (PdfObject child in array.Items) pending.Push(child);
            }
        }
        return result;
    }

    private static OfficeProvenanceOptions CreateOutputInspectionOptions(OfficeProvenanceOptions source, long outputBytes) => new() {
        MaxAssetBytes = Math.Max(source.MaxAssetBytes, outputBytes),
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = source.MaxEmbeddedAssets
    };

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
            if (!ReferenceEquals(dictionary, catalog) && IsInformationResource(objects, value, dictionary, structuralAssociationSites)) {
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
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject owner,
        PdfDictionary dictionary,
        HashSet<PdfObject> structuralAssociationSites) {
        if (structuralAssociationSites.Contains(dictionary)) return false;
        string? type = PdfObjectLookup.Resolve(
            objects,
            dictionary.Items.TryGetValue("Type", out PdfObject? typeValue) ? typeValue : null) is PdfName typeName
                ? typeName.Name
                : null;
        string? subtype = PdfObjectLookup.Resolve(
            objects,
            dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeValue) ? subtypeValue : null) is PdfName subtypeName
                ? subtypeName.Name
                : null;
        if (type is "Catalog" or "Pages" or "Page" or "Annot" or "Filespec" or "EmbeddedFile" or "XRef" or "ObjStm") return false;
        if (subtype is "FileAttachment" or "Popup" ||
            subtype != null && dictionary.Items.ContainsKey("Rect")) return false;
        if (dictionary.Items.ContainsKey("EF") &&
            (dictionary.Items.ContainsKey("F") || dictionary.Items.ContainsKey("UF"))) return false;
        if (owner is PdfStream) return true;
        return string.Equals(type, "StructElem", StringComparison.Ordinal) || type == null;
    }

    private static HashSet<PdfObject> CollectStructuralAssociationSites(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        PdfReference? activeInfoReference,
        int? activeEncryptObjectNumber,
        HashSet<int> reachableObjectNumbers,
        HashSet<int> pageTreeObjectNumbers,
        HashSet<int> activePageObjectNumbers,
        int maximumContainerEntries) {
        var result = new HashSet<PdfObject>();
        AddResolvedDictionary(objects, activeInfoReference, result);
        if (activeEncryptObjectNumber.HasValue &&
            objects.TryGetValue(activeEncryptObjectNumber.Value, out PdfIndirectObject? activeEncryption)) {
            AddResolvedDictionary(objects, activeEncryption.Value, result);
        }
        foreach (int objectNumber in pageTreeObjectNumbers) {
            if (objects.TryGetValue(objectNumber, out PdfIndirectObject? pageTreeObject)) {
                AddResolvedDictionary(objects, pageTreeObject.Value, result);
            }
        }
        foreach (string key in new[] { "AcroForm", "ViewerPreferences", "MarkInfo", "StructTreeRoot" }) {
            AddResolvedDictionary(objects, catalog.Items.TryGetValue(key, out PdfObject? value) ? value : null, result);
        }
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Perms", out PdfObject? permissions) ? permissions : null,
            result,
            maximumContainerEntries);
        if (PdfObjectLookup.Resolve(objects,
                catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? structureTreeValue) ? structureTreeValue : null) is PdfDictionary structureTree) {
            AddStructuralGraphDictionaries(
                objects,
                structureTree.Items.TryGetValue("ParentTree", out PdfObject? parentTree) ? parentTree : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers);
            AddStructuralGraphDictionaries(
                objects,
                structureTree.Items.TryGetValue("RoleMap", out PdfObject? roleMap) ? roleMap : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers);
            AddStructuralGraphDictionaries(
                objects,
                structureTree.Items.TryGetValue("ClassMap", out PdfObject? classMap) ? classMap : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers);
            AddStructuralGraphDictionaries(
                objects,
                structureTree.Items.TryGetValue("IDTree", out PdfObject? idTree) ? idTree : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers);
        }
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Extensions", out PdfObject? extensions) ? extensions : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContent) ? optionalContent : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Collection", out PdfObject? collection) ? collection : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("URI", out PdfObject? uri) ? uri : null,
            result,
            maximumContainerEntries);
        foreach (string key in new[] { "OpenAction", "AA" }) {
            AddStructuralGraphDictionaries(
                objects,
                catalog.Items.TryGetValue(key, out PdfObject? action) ? action : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers);
        }
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Dests", out PdfObject? destinations) ? destinations : null,
            result,
            maximumContainerEntries,
            pageTreeObjectNumbers);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("OutputIntents", out PdfObject? outputIntents) ? outputIntents : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("DSS", out PdfObject? documentSecurityStore) ? documentSecurityStore : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Legal", out PdfObject? legal) ? legal : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("PieceInfo", out PdfObject? catalogPieceInfo) ? catalogPieceInfo : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("SpiderInfo", out PdfObject? spiderInfo) ? spiderInfo : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Requirements", out PdfObject? requirements) ? requirements : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("Threads", out PdfObject? threads) ? threads : null,
            result,
            maximumContainerEntries,
            pageTreeObjectNumbers);
        AddAcroFormFieldDictionaries(
            objects,
            catalog.Items.TryGetValue("AcroForm", out PdfObject? acroForm) ? acroForm : null,
            result,
            maximumContainerEntries);
        AddOutlineDictionaries(
            objects,
            catalog.Items.TryGetValue("Outlines", out PdfObject? outlines) ? outlines : null,
            result,
            maximumContainerEntries,
            pageTreeObjectNumbers);
        AddCatalogNameTrees(
            objects,
            catalog.Items.TryGetValue("Names", out PdfObject? names) ? names : null,
            result,
            maximumContainerEntries,
            pageTreeObjectNumbers);
        AddStructuralGraphDictionaries(
            objects,
            catalog.Items.TryGetValue("PageLabels", out PdfObject? pageLabels) ? pageLabels : null,
            result,
            maximumContainerEntries);
        AddEmbeddedFileGraphDictionaries(objects, reachableObjectNumbers, result);
        var resourceSites = new HashSet<PdfObject>();
        var structuralTraversalVisited = new HashSet<PdfObject>();
        var annotationStructuralVisited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values.Where(item => reachableObjectNumbers.Contains(item.ObjectNumber))) {
            PdfDictionary? dictionary = item.Value is PdfStream stream ? stream.Dictionary : item.Value as PdfDictionary;
            if (dictionary == null) continue;
            if (activePageObjectNumbers.Contains(item.ObjectNumber)) {
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("PieceInfo", out PdfObject? pieceInfo) ? pieceInfo : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("Group", out PdfObject? group) ? group : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("Trans", out PdfObject? transition) ? transition : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("VP", out PdfObject? viewports) ? viewports : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("BoxColorInfo", out PdfObject? boxColorInfo) ? boxColorInfo : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("SeparationInfo", out PdfObject? separationInfo) ? separationInfo : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue("PresSteps", out PdfObject? presentationSteps) ? presentationSteps : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
                AddPageAnnotationDictionaries(objects, dictionary, result, maximumContainerEntries, annotationStructuralVisited);
            }
            if (item.Value is PdfStream activeStream) {
                foreach (string key in new[] { "DecodeParms", "DP", "FDecodeParms" }) {
                    AddStructuralGraphDictionaries(
                        objects,
                        activeStream.Dictionary.Items.TryGetValue(key, out PdfObject? decodingParameters) ? decodingParameters : null,
                        result,
                        maximumContainerEntries,
                        sharedVisited: structuralTraversalVisited);
                }
                string? streamSubtype = GetResolvedName(objects, activeStream.Dictionary, "Subtype");
                if (string.Equals(streamSubtype, "Form", StringComparison.Ordinal) ||
                    string.Equals(streamSubtype, "Image", StringComparison.Ordinal)) {
                    AddStructuralGraphDictionaries(
                        objects,
                        activeStream.Dictionary.Items.TryGetValue("OC", out PdfObject? xObjectOptionalContent) ? xObjectOptionalContent : null,
                        result,
                        maximumContainerEntries,
                        sharedVisited: structuralTraversalVisited);
                }
                if (string.Equals(streamSubtype, "Form", StringComparison.Ordinal)) {
                    foreach (string key in new[] { "Group", "Ref", "PieceInfo" }) {
                        AddStructuralGraphDictionaries(
                            objects,
                            activeStream.Dictionary.Items.TryGetValue(key, out PdfObject? formValue) ? formValue : null,
                            result,
                            maximumContainerEntries,
                            sharedVisited: structuralTraversalVisited);
                    }
                } else if (string.Equals(streamSubtype, "Image", StringComparison.Ordinal)) {
                    foreach (string key in new[] { "ColorSpace", "Alternates", "OPI" }) {
                        AddStructuralGraphDictionaries(
                            objects,
                            activeStream.Dictionary.Items.TryGetValue(key, out PdfObject? imageValue) ? imageValue : null,
                            result,
                            maximumContainerEntries,
                            sharedVisited: structuralTraversalVisited);
                    }
                }
                if (activeStream.Dictionary.Items.ContainsKey("ShadingType")) {
                    foreach (string key in new[] { "Function", "ColorSpace" }) {
                        AddStructuralGraphDictionaries(
                            objects,
                            activeStream.Dictionary.Items.TryGetValue(key, out PdfObject? shadingValue) ? shadingValue : null,
                            result,
                            maximumContainerEntries,
                            sharedVisited: structuralTraversalVisited);
                    }
                }
                AddStructuralGraphDictionaries(
                    objects,
                    activeStream.Dictionary.Items.TryGetValue("F", out PdfObject? externalFile) ? externalFile : null,
                    result,
                    maximumContainerEntries,
                    sharedVisited: structuralTraversalVisited);
            }
            PdfObject? resources = dictionary.Items.TryGetValue("Resources", out PdfObject? resourceValue) ? resourceValue : null;
            PdfObject? defaultResources = dictionary.Items.TryGetValue("DR", out PdfObject? defaultResourceValue) ? defaultResourceValue : null;
            AddResourceDictionaries(objects, resources, result, resourceSites);
            AddResourceDictionaries(objects, defaultResources, result, resourceSites);
            AddIccBasedAlternateDictionaries(objects, resources, result, maximumContainerEntries, structuralTraversalVisited);
            AddIccBasedAlternateDictionaries(objects, defaultResources, result, maximumContainerEntries, structuralTraversalVisited);
            if (string.Equals(GetResolvedName(objects, dictionary, "Type"), "EmbeddedFile", StringComparison.Ordinal)) {
                AddResolvedDictionary(objects, dictionary.Items.TryGetValue("Params", out PdfObject? parameters) ? parameters : null, result);
            }
        }
        AddFileSpecificationDescendantGraphs(objects, reachableObjectNumbers, result, maximumContainerEntries);
        foreach (PdfDictionary owner in result.ToArray()) {
            foreach (string key in new[] { "A", "AA", "OpenAction" }) {
                AddStructuralGraphDictionaries(
                    objects,
                    owner.Items.TryGetValue(key, out PdfObject? action) ? action : null,
                    result,
                    maximumContainerEntries,
                    pageTreeObjectNumbers,
                    structuralTraversalVisited);
            }
        }
        return result;
    }

    private static void AddStructuralGraphDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        HashSet<int>? terminalObjectNumbers = null,
        HashSet<PdfObject>? sharedVisited = null) {
        if (value == null) return;
        HashSet<PdfObject> visited = sharedVisited ?? new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(value);
        while (pending.Count > 0) {
            PdfObject current = pending.Pop();
            if (current is PdfReference reference && terminalObjectNumbers?.Contains(reference.ObjectNumber) == true) continue;
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, current);
            if (resolved == null || !visited.Add(resolved)) continue;
            if (visited.Count > maximumContainerEntries) {
                throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
            }
            PdfDictionary? dictionary = resolved is PdfStream stream ? stream.Dictionary : resolved as PdfDictionary;
            if (dictionary != null) {
                result.Add(dictionary);
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            } else if (resolved is PdfArray array) {
                foreach (PdfObject child in array.Items) pending.Push(child);
            }
        }
    }

    private static void AddAcroFormFieldDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result,
        int maximumContainerEntries) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfDictionary acroForm) {
            return;
        }
        AddStructuralGraphDictionaries(
            objects,
            acroForm.Items.TryGetValue("DR", out PdfObject? defaultResources) ? defaultResources : null,
            result,
            maximumContainerEntries);
        AddStructuralGraphDictionaries(
            objects,
            acroForm.Items.TryGetValue("XFA", out PdfObject? xfaValue) ? xfaValue : null,
            result,
            maximumContainerEntries);
        if (PdfObjectLookup.Resolve(objects, acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsValue) ? fieldsValue : null) is not PdfArray fields) return;
        var visited = new HashSet<PdfObject>();
        var structuralVisited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>(fields.Items);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary field || !visited.Add(field)) continue;
            if (visited.Count > maximumContainerEntries) {
                throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
            }
            result.Add(field);
            foreach (string key in new[] { "Lock", "SV", "V", "AP", "MK", "BS" }) {
                AddStructuralGraphDictionaries(
                    objects,
                    field.Items.TryGetValue(key, out PdfObject? constraintValue) ? constraintValue : null,
                    result,
                    maximumContainerEntries,
                    sharedVisited: structuralVisited);
            }
            if (PdfObjectLookup.Resolve(objects, field.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) continue;
            foreach (PdfObject child in kids.Items) pending.Push(child);
        }
    }

    private static void AddOutlineDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        HashSet<int> pageTreeObjectNumbers) {
        if (value == null) return;
        var visited = new HashSet<PdfObject>();
        var structuralVisited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(value);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary outline || !visited.Add(outline)) continue;
            if (visited.Count > maximumContainerEntries) {
                throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
            }
            result.Add(outline);
            AddStructuralGraphDictionaries(
                objects,
                outline.Items.TryGetValue("Dest", out PdfObject? destination) ? destination : null,
                result,
                maximumContainerEntries,
                pageTreeObjectNumbers,
                structuralVisited);
            foreach (string key in new[] { "First", "Last", "Next", "Prev", "Parent" }) {
                if (outline.Items.TryGetValue(key, out PdfObject? linked)) pending.Push(linked);
            }
        }
    }

    private static void AddEmbeddedFileGraphDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> reachableObjectNumbers,
        HashSet<PdfObject> result) {
        foreach (PdfIndirectObject item in objects.Values.Where(item => reachableObjectNumbers.Contains(item.ObjectNumber))) {
            if (!IsFileSpecificationValue(objects, item.Value) || item.Value is not PdfDictionary fileSpecification ||
                PdfObjectLookup.Resolve(objects, fileSpecification.Items.TryGetValue("EF", out PdfObject? embeddedFilesValue) ? embeddedFilesValue : null) is not PdfDictionary embeddedFiles) continue;
            result.Add(embeddedFiles);
            foreach (PdfObject variant in embeddedFiles.Items.Values) {
                if (PdfObjectLookup.Resolve(objects, variant) is not PdfStream embeddedFile) continue;
                result.Add(embeddedFile.Dictionary);
                AddResolvedDictionary(
                    objects,
                    embeddedFile.Dictionary.Items.TryGetValue("Params", out PdfObject? parameters) ? parameters : null,
                    result);
            }
        }
    }

    private static void AddResolvedDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        PdfDictionary? dictionary = resolved is PdfStream stream ? stream.Dictionary : resolved as PdfDictionary;
        if (dictionary != null) result.Add(dictionary);
    }

    private static void AddCatalogNameTrees(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        HashSet<int> pageTreeObjectNumbers) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfDictionary names) return;
        result.Add(names);
        var ordinaryVisited = new HashSet<PdfObject>();
        var destinationVisited = new HashSet<PdfObject>();
        var actionVisited = new HashSet<PdfObject>();
        var ordinaryLeafVisited = new HashSet<PdfObject>();
        var destinationLeafVisited = new HashSet<PdfObject>();
        var actionLeafVisited = new HashSet<PdfObject>();
        foreach (KeyValuePair<string, PdfObject> item in names.Items) {
            bool addDestinationLeafValues = string.Equals(item.Key, "Dests", StringComparison.Ordinal);
            bool addActionLeafGraphs = string.Equals(item.Key, "JavaScript", StringComparison.Ordinal);
            bool addOrdinaryLeafGraphs = !string.Equals(item.Key, "EmbeddedFiles", StringComparison.Ordinal) &&
                                         !addDestinationLeafValues &&
                                         !addActionLeafGraphs;
            AddNameTreeDictionaries(
                objects,
                new[] { item.Value },
                result,
                maximumContainerEntries,
                addDestinationLeafValues,
                addActionLeafGraphs,
                addOrdinaryLeafGraphs,
                pageTreeObjectNumbers,
                addDestinationLeafValues ? destinationVisited : addActionLeafGraphs ? actionVisited : ordinaryVisited,
                addDestinationLeafValues ? destinationLeafVisited : addActionLeafGraphs ? actionLeafVisited : ordinaryLeafVisited);
        }
    }

    private static void AddNameTreeDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<PdfObject?> values,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        bool addDestinationLeafValues,
        bool addActionLeafGraphs,
        bool addOrdinaryLeafGraphs,
        HashSet<int> pageTreeObjectNumbers,
        HashSet<PdfObject> visited,
        HashSet<PdfObject> leafStructuralVisited) {
        var pending = new Stack<PdfObject>(values.Where(static value => value != null).Cast<PdfObject>());
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary dictionary || !visited.Add(dictionary)) continue;
            if (visited.Count > maximumContainerEntries) {
                throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
            }
            result.Add(dictionary);
            if ((addDestinationLeafValues || addActionLeafGraphs || addOrdinaryLeafGraphs) && PdfObjectLookup.Resolve(objects,
                    dictionary.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is PdfArray leafNames) {
                for (int index = 1; index < leafNames.Items.Count; index += 2) {
                    if (addActionLeafGraphs || addOrdinaryLeafGraphs) {
                        AddStructuralGraphDictionaries(
                            objects,
                            leafNames.Items[index],
                            result,
                            maximumContainerEntries,
                            pageTreeObjectNumbers,
                            leafStructuralVisited);
                    } else {
                        AddResolvedDictionary(objects, leafNames.Items[index], result);
                    }
                }
            }
            if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) continue;
            foreach (PdfObject child in kids.Items) pending.Push(child);
        }
    }

    private static void AddPageAnnotationDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary page,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        HashSet<PdfObject> structuralVisited) {
        if (PdfObjectLookup.Resolve(objects,
                page.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray annotations) return;
        if (annotations.Items.Count > maximumContainerEntries) {
            throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
        }
        foreach (PdfObject annotation in annotations.Items) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, annotation);
            PdfDictionary? dictionary = resolved is PdfStream stream ? stream.Dictionary : resolved as PdfDictionary;
            if (dictionary == null) continue;
            result.Add(dictionary);
            foreach (string key in new[] { "AP", "BS", "BE", "MK", "Dest", "Movie", "3DV", "RichMediaContent", "RichMediaSettings", "FixedPrint" }) {
                AddStructuralGraphDictionaries(
                    objects,
                    dictionary.Items.TryGetValue(key, out PdfObject? structuralValue) ? structuralValue : null,
                    result,
                    maximumContainerEntries,
                    sharedVisited: structuralVisited);
            }
        }
    }

    private static void AddFileSpecificationDescendantGraphs(
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> reachableObjectNumbers,
        HashSet<PdfObject> result,
        int maximumContainerEntries) {
        var structuralVisited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values.Where(item => reachableObjectNumbers.Contains(item.ObjectNumber))) {
            if (!IsFileSpecificationValue(objects, item.Value) || item.Value is not PdfDictionary fileSpecification) continue;
            foreach (PdfObject child in fileSpecification.Items.Values) {
                AddStructuralGraphDictionaries(
                    objects,
                    child,
                    result,
                    maximumContainerEntries,
                    sharedVisited: structuralVisited);
            }
        }
    }

    private static void AddResourceDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<PdfObject> result,
        HashSet<PdfObject> visited) {
        if (value == null) return;
        var pending = new Stack<PdfObject>();
        pending.Push(value);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved == null || resolved is PdfStream || !visited.Add(resolved)) continue;
            if (resolved is PdfDictionary dictionary) {
                result.Add(dictionary);
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            } else if (resolved is PdfArray array) {
                foreach (PdfObject child in array.Items) pending.Push(child);
            }
        }
    }

    private static void AddIccBasedAlternateDictionaries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? resourcesValue,
        HashSet<PdfObject> result,
        int maximumContainerEntries,
        HashSet<PdfObject> structuralVisited) {
        if (PdfObjectLookup.Resolve(objects, resourcesValue) is not PdfDictionary resources ||
            PdfObjectLookup.Resolve(objects,
                resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesValue) ? colorSpacesValue : null) is not PdfDictionary colorSpaces) {
            return;
        }
        foreach (PdfObject colorSpaceValue in colorSpaces.Items.Values) {
            if (PdfObjectLookup.Resolve(objects, colorSpaceValue) is not PdfArray colorSpace || colorSpace.Items.Count < 2 ||
                PdfObjectLookup.Resolve(objects, colorSpace.Items[0]) is not PdfName family ||
                !string.Equals(family.Name, "ICCBased", StringComparison.Ordinal) ||
                PdfObjectLookup.Resolve(objects, colorSpace.Items[1]) is not PdfStream profile) {
                continue;
            }
            AddStructuralGraphDictionaries(
                objects,
                profile.Dictionary.Items.TryGetValue("Alternate", out PdfObject? alternate) ? alternate : null,
                result,
                maximumContainerEntries,
                sharedVisited: structuralVisited);
        }
    }

    private static void CollectPageAnnotationReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        HashSet<int> result,
        HashSet<int> structuralObjectNumbers,
        int maximumAnnotationsPerPage) {
        if (!catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return;
        var pending = new Stack<PdfObject>();
        var visited = new HashSet<PdfObject>();
        pending.Push(pages);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary dictionary || !visited.Add(resolved)) continue;
            string? type = GetResolvedName(objects, dictionary, "Type");
            PdfArray? kids = PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) as PdfArray;
            if (type == "Pages" || kids != null) {
                if (kids != null) {
                    foreach (PdfObject child in kids.Items) pending.Push(child);
                }
                continue;
            }
            if ((type != null && type != "Page") ||
                PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray annotations) continue;
            if (annotations.Items.Count > maximumAnnotationsPerPage) {
                throw new InvalidDataException("PDF page annotations exceed the configured per-page limit.");
            }
            foreach (PdfObject annotationValue in annotations.Items) {
                if (annotationValue is PdfReference annotationReference &&
                    PdfObjectLookup.TryGet(objects, annotationReference, out _)) {
                    structuralObjectNumbers.Add(annotationReference.ObjectNumber);
                }
                if (PdfObjectLookup.Resolve(objects, annotationValue) is not PdfDictionary annotation ||
                    !string.Equals(GetResolvedName(objects, annotation, "Subtype"), "FileAttachment", StringComparison.Ordinal) ||
                    !annotation.Items.TryGetValue("FS", out PdfObject? fileSpecification) || fileSpecification is not PdfReference reference ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    !IsFileSpecificationValue(objects, indirect.Value)) continue;
                result.Add(reference.ObjectNumber);
            }
        }
    }

    private static bool IsFileSpecificationObject(
        Dictionary<int, PdfIndirectObject> objects,
        int objectNumber,
        HashSet<int> pageTreeObjectNumbers,
        HashSet<int> structuralObjectNumbers) =>
        !pageTreeObjectNumbers.Contains(objectNumber) &&
        !structuralObjectNumbers.Contains(objectNumber) &&
        objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) &&
        IsFileSpecificationValue(objects, indirect.Value);

    private static bool HasOnlySelectedEmbeddedFileVariants(
        Dictionary<int, PdfIndirectObject> objects,
        PdfExtractedAttachment attachment) {
        if (attachment.EmbeddedFileObjectNumber <= 0 ||
            !objects.TryGetValue(attachment.FileSpecObjectNumber, out PdfIndirectObject? fileSpecObject) ||
            fileSpecObject.Value is not PdfDictionary fileSpec ||
            PdfObjectLookup.Resolve(objects, fileSpec.Items.TryGetValue("EF", out PdfObject? embeddedFilesValue) ? embeddedFilesValue : null) is not PdfDictionary embeddedFiles ||
            !objects.TryGetValue(attachment.EmbeddedFileObjectNumber, out PdfIndirectObject? selectedEmbeddedFile) ||
            selectedEmbeddedFile.Value is not PdfStream selectedStream ||
            embeddedFiles.Items.Count == 0) return false;
        foreach (PdfObject variant in embeddedFiles.Items.Values) {
            if (!ReferenceEquals(PdfObjectLookup.Resolve(objects, variant), selectedStream)) return false;
        }
        return true;
    }

    private static bool HasEmbeddedFileStreamType(
        Dictionary<int, PdfIndirectObject> objects,
        PdfExtractedAttachment attachment) =>
        objects.TryGetValue(attachment.EmbeddedFileObjectNumber, out PdfIndirectObject? embeddedFile) &&
        embeddedFile.Value is PdfStream stream &&
        string.Equals(GetResolvedName(objects, stream.Dictionary, "Type"), "EmbeddedFile", StringComparison.Ordinal);

    private static string? GetResolvedName(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key) =>
        PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null) is PdfName name
            ? name.Name
            : null;

    private static HashSet<int> CollectPageTreeObjectNumbers(PdfReadDocument document, int maximumContainerEntries) {
        var result = new HashSet<int>(document.Pages.Select(static page => page.ObjectNumber));
        PdfDictionary? catalog = PdfSyntax.FindCatalog(document.Objects, document.TrailerRaw);
        if (catalog == null || !catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return result;
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(pages);
        while (pending.Count > 0) {
            PdfObject value = pending.Pop();
            if (value is PdfReference reference) {
                if (!PdfObjectLookup.TryGet(document.Objects, reference, out PdfIndirectObject? indirect) ||
                    !result.Add(reference.ObjectNumber)) continue;
                if (result.Count > maximumContainerEntries) {
                    throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
                }
                value = indirect.Value;
            }
            PdfObject? resolved = PdfObjectLookup.Resolve(document.Objects, value);
            if (resolved is not PdfDictionary dictionary || !visited.Add(dictionary) ||
                PdfObjectLookup.Resolve(document.Objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) continue;
            foreach (PdfObject child in kids.Items) pending.Push(child);
        }
        return result;
    }

    private static bool IsFileSpecificationValue(Dictionary<int, PdfIndirectObject> objects, PdfObject value) {
        if (value is not PdfDictionary dictionary) return false;
        bool hasType = dictionary.Items.TryGetValue("Type", out PdfObject? typeValue);
        PdfObject? resolvedType = PdfObjectLookup.Resolve(objects, typeValue);
        if (hasType && resolvedType is not PdfName) return false;
        string? type = (resolvedType as PdfName)?.Name;
        if (string.Equals(type, "Annot", StringComparison.Ordinal) || dictionary.Items.ContainsKey("Subtype")) return false;
        bool hasFileName = PdfObjectLookup.Resolve(objects,
                dictionary.Items.TryGetValue("UF", out PdfObject? unicodeName) ? unicodeName : null) is PdfStringObj unicodeText &&
                !string.IsNullOrEmpty(unicodeText.Value) ||
            PdfObjectLookup.Resolve(objects,
                dictionary.Items.TryGetValue("F", out PdfObject? fileName) ? fileName : null) is PdfStringObj fileText &&
                !string.IsNullOrEmpty(fileText.Value);
        return hasFileName && (type == null || string.Equals(type, "Filespec", StringComparison.Ordinal));
    }

    private static void AddReferencesFromArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<int> result) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfArray array) return;
        foreach (PdfObject item in array.Items) {
            if (item is PdfReference reference &&
                PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                IsFileSpecificationValue(objects, indirect.Value)) result.Add(reference.ObjectNumber);
        }
    }

    private static void CollectEmbeddedFilesNameTreeReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        HashSet<int> result,
        int maximumContainerEntries) {
        if (PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is not PdfDictionary names ||
            !names.Items.TryGetValue("EmbeddedFiles", out PdfObject? embeddedFiles)) return;
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(embeddedFiles);
        while (pending.Count > 0) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (resolved is not PdfDictionary dictionary || !visited.Add(dictionary)) continue;
            if (visited.Count > maximumContainerEntries) {
                throw new InvalidDataException($"The PDF exceeds the configured container entry limit of {maximumContainerEntries}.");
            }
            if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Names", out PdfObject? leafNamesValue) ? leafNamesValue : null) is PdfArray leafNames) {
                for (int index = 1; index < leafNames.Items.Count; index += 2) {
                    if (PdfObjectLookup.Resolve(objects, leafNames.Items[index - 1]) is PdfStringObj &&
                        leafNames.Items[index] is PdfReference reference &&
                        PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                        IsFileSpecificationValue(objects, indirect.Value)) result.Add(reference.ObjectNumber);
                }
            }
            if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) continue;
            foreach (PdfObject child in kids.Items) pending.Push(child);
        }
    }

    private sealed class PdfC2paAssociationProfile {
        private readonly HashSet<int> _documentLevel;
        private readonly HashSet<int> _objectLevel;
        private readonly HashSet<int> _secondaryDocumentReferences;
        private readonly HashSet<int> _structuralObjectNumbers;

        internal PdfC2paAssociationProfile(HashSet<int> documentLevel, HashSet<int> objectLevel, HashSet<int> secondaryDocumentReferences, HashSet<int> structuralObjectNumbers) {
            _documentLevel = documentLevel;
            _objectLevel = objectLevel;
            _secondaryDocumentReferences = secondaryDocumentReferences;
            _structuralObjectNumbers = structuralObjectNumbers;
        }

        internal HashSet<int> StructuralObjectNumbers => _structuralObjectNumbers;

        internal bool IsValid(int fileSpecObjectNumber) => fileSpecObjectNumber > 0 &&
            (_objectLevel.Contains(fileSpecObjectNumber) ||
             _documentLevel.Contains(fileSpecObjectNumber) && _secondaryDocumentReferences.Contains(fileSpecObjectNumber));
    }

    private static byte[] ReadBounded(string filePath, long maximumBytes) {
        using var stream = File.OpenRead(Path.GetFullPath(filePath));
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes);
    }
}
