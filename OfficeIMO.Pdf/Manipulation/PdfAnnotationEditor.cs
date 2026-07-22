using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

/// <summary>Edits or removes PDF annotations without third-party dependencies.</summary>
internal static partial class PdfAnnotationEditor {
    private static readonly HashSet<string> KnownAnnotationSubtypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "Text",
        "Link",
        "FreeText",
        "Line",
        "Square",
        "Circle",
        "Polygon",
        "PolyLine",
        "Highlight",
        "Underline",
        "Squiggly",
        "StrikeOut",
        "Caret",
        "Stamp",
        "Ink",
        "Popup",
        "FileAttachment",
        "Sound",
        "Movie",
        "Widget",
        "Screen",
        "PrinterMark",
        "TrapNet",
        "Watermark",
        "3D",
        "Redact"
    };

    /// <summary>Removes annotations matching the supplied filters and returns rewritten PDF bytes.</summary>
    public static PdfAnnotationEditResult RemoveAnnotations(byte[] pdf, PdfAnnotationRemovalOptions? options = null) => RemoveAnnotations(pdf, options, readOptions: null);

    /// <summary>Removes annotations using explicit read limits or credentials and returns rewritten PDF bytes.</summary>
    public static PdfAnnotationEditResult RemoveAnnotations(byte[] pdf, PdfAnnotationRemovalOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));

        PdfAnnotationRemovalOptions effectiveOptions = options ?? new PdfAnnotationRemovalOptions();
        ValidateRemovalOptions(effectiveOptions);
        PdfMutationPlan mutationPlan = PdfMutationPlanner.Require(
            pdf,
            PdfMutationOperation.ModifyAnnotations,
            readOptions,
            executionPreference: effectiveOptions.ExecutionPreference);
        if (mutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            if (!effectiveOptions.AllowResidualDataInAppendOnly) {
                throw new NotSupportedException("Append-only annotation removal retains the original annotation bytes in prior revisions. Use a permitted full rewrite for sanitization or explicitly allow residual data.");
            }
            return RemoveAnnotationsIncrementally(pdf, effectiveOptions, mutationPlan, readOptions);
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        List<int> pageObjectNumbers = GetPageObjectNumbersInDocumentOrder(objects);
        var removed = new HashSet<int>();
        bool changed = false;
        for (int pageIndex = 0; pageIndex < pageObjectNumbers.Count; pageIndex++) {
            if (effectiveOptions.PageNumber.HasValue && effectiveOptions.PageNumber.Value != pageIndex + 1) {
                continue;
            }

            if (!objects.TryGetValue(pageObjectNumbers[pageIndex], out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary page ||
                !page.Items.TryGetValue("Annots", out PdfObject? annotsObject)) {
                continue;
            }

            PdfObject resolvedAnnots = PdfObjectLookup.Resolve(objects, annotsObject) ?? annotsObject;
            if (resolvedAnnots is not PdfArray annotations) {
                continue;
            }

            for (int i = annotations.Items.Count - 1; i >= 0; i--) {
                PdfObject item = annotations.Items[i];
                int? annotationObjectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
                PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
                if (annotation is null || !MatchesRemovalFilter(objects, annotation, annotationObjectNumber, effectiveOptions)) {
                    continue;
                }

                if (effectiveOptions.RemoveMatchingPopups &&
                    annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
                    popupObject is PdfReference popupReference) {
                    removed.Add(popupReference.ObjectNumber);
                }

                if (annotationObjectNumber.HasValue) {
                    removed.Add(annotationObjectNumber.Value);
                }

                annotations.Items.RemoveAt(i);
                changed = true;
            }

            if (effectiveOptions.RemoveMatchingPopups) {
                RemovePopupReferences(objects, annotations, removed, ref changed);
            }

            if (annotations.Items.Count == 0) {
                page.Items.Remove("Annots");
                changed = true;
            }
        }

        if (removed.Count > 0) {
            ClearRemovedPopupReferences(objects, removed);
        }

        foreach (int objectNumber in removed) {
            objects.Remove(objectNumber);
        }

        if (!changed && removed.Count == 0) {
            return CreateFullRewriteResult(pdf, (byte[])pdf.Clone(), 0, mutationPlan, annotationsChanged: false, readOptions: readOptions);
        }

        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata, pdf);
        return CreateFullRewriteResult(pdf, rewritten, Math.Max(removed.Count, 1), mutationPlan, annotationsChanged: true, readOptions: readOptions);
    }

    /// <summary>Updates a single indirect annotation and returns rewritten PDF bytes.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(byte[] pdf, int objectNumber, PdfAnnotationUpdateOptions options) => UpdateAnnotation(pdf, objectNumber, options, readOptions: null);

    /// <summary>Updates one indirect annotation using explicit read limits or credentials.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(byte[] pdf, int objectNumber, PdfAnnotationUpdateOptions options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        if (objectNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "Annotation object number must be positive.");
        }

        ValidateUpdateOptions(options);
        PdfMutationPlan mutationPlan = PdfMutationPlanner.Require(
            pdf,
            PdfMutationOperation.ModifyAnnotations,
            readOptions,
            executionPreference: options.ExecutionPreference);
        if (mutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            if (!options.AllowResidualDataInAppendOnly) {
                throw new NotSupportedException("Append-only annotation updates retain replaced annotation data in prior revisions. Use a permitted full rewrite for sanitization or explicitly allow residual data.");
            }
            return UpdateAnnotationIncrementally(pdf, objectNumber, options, mutationPlan, readOptions);
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfDictionary annotation ||
            !IsAnnotationUpdateTarget(objects, objectNumber, annotation)) {
            throw new ArgumentException("PDF annotation object was not found: " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(objectNumber));
        }

        IReadOnlyList<int> changedObjects = ApplyUpdates(objects, annotation, options);
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata, pdf);
        return CreateFullRewriteResult(pdf, rewritten, 1, mutationPlan, annotationsChanged: false, readOptions: readOptions);
    }

    /// <summary>Removes annotations from a PDF file and writes the result to another file.</summary>
    public static PdfAnnotationEditResult RemoveAnnotations(string inputPath, string outputPath, PdfAnnotationRemovalOptions? options = null) => RemoveAnnotations(inputPath, outputPath, options, readOptions: null);

    /// <summary>Removes annotations from a PDF file using explicit read limits or credentials and writes the result to another file.</summary>
    public static PdfAnnotationEditResult RemoveAnnotations(string inputPath, string outputPath, PdfAnnotationRemovalOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        PdfAnnotationEditResult result = RemoveAnnotations(File.ReadAllBytes(inputPath), options, readOptions);
        WriteFile(fullOutputPath, result.Bytes);
        return result;
    }

    /// <summary>Updates a single annotation in a PDF file and writes the result to another file.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(string inputPath, string outputPath, int objectNumber, PdfAnnotationUpdateOptions options) => UpdateAnnotation(inputPath, outputPath, objectNumber, options, readOptions: null);

    /// <summary>Updates one annotation in a PDF file using explicit read limits or credentials and writes the result to another file.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(string inputPath, string outputPath, int objectNumber, PdfAnnotationUpdateOptions options, PdfReadOptions? readOptions) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        PdfAnnotationEditResult result = UpdateAnnotation(File.ReadAllBytes(inputPath), objectNumber, options, readOptions);
        WriteFile(fullOutputPath, result.Bytes);
        return result;
    }

    private static bool MatchesRemovalFilter(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, int? objectNumber, PdfAnnotationRemovalOptions options) {
        if (options.ObjectNumber.HasValue && objectNumber != options.ObjectNumber.Value) {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(options.Subtype)) {
            string? subtype = TryReadName(objects, annotation, "Subtype");
            if (!string.Equals(subtype, options.Subtype, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }
        }

        return IsAnnotation(annotation) || HasAnnotationSubtype(annotation);
    }

    private static int RemovePageAnnotationReferences(PdfArray annotations, HashSet<int> removedObjectNumbers) {
        int removed = 0;
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            if (annotations.Items[i] is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber)) {
                annotations.Items.RemoveAt(i);
                removed++;
            }
        }

        return removed;
    }

    private static void ClearRemovedPopupReferences(Dictionary<int, PdfIndirectObject> objects, HashSet<int> removedObjectNumbers) {
        if (removedObjectNumbers.Count == 0) {
            return;
        }

        foreach (PdfIndirectObject indirect in objects.Values) {
            ClearRemovedPopupReferences(indirect.Value, removedObjectNumbers, new HashSet<PdfObject>());
        }
    }

    private static void ClearRemovedPopupReferences(PdfObject value, HashSet<int> removedObjectNumbers, HashSet<PdfObject> visited) {
        if (!visited.Add(value)) {
            return;
        }

        if (value is PdfDictionary dictionary) {
            if (dictionary.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
                popupObject is PdfReference popupReference &&
                removedObjectNumbers.Contains(popupReference.ObjectNumber)) {
                dictionary.Items.Remove("Popup");
            }

            foreach (PdfObject child in dictionary.Items.Values.ToArray()) {
                ClearRemovedPopupReferences(child, removedObjectNumbers, visited);
            }

            return;
        }

        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items.ToArray()) {
                ClearRemovedPopupReferences(child, removedObjectNumbers, visited);
            }
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<int> ApplyUpdates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, PdfAnnotationUpdateOptions options) {
        var changedObjects = new List<int>();
        bool invalidateAppearance = false;
        if (options.Contents is not null) {
            annotation.Items["Contents"] = new PdfStringObj(options.Contents, useTextStringEncoding: true);
            invalidateAppearance = true;
        }

        if (options.Title is not null) {
            annotation.Items["T"] = new PdfStringObj(options.Title, useTextStringEncoding: true);
        }

        if (options.Name is not null) {
            annotation.Items["NM"] = new PdfStringObj(options.Name, useTextStringEncoding: true);
        }

        if (options.Flags.HasValue) {
            annotation.Items["F"] = new PdfNumber(options.Flags.Value);
        }

        if (options.Color is not null) {
            annotation.Items["C"] = CreateColorArray(options.Color);
            invalidateAppearance = true;
        }

        if (options.RemoveActions) {
            annotation.Items.Remove("A");
            annotation.Items.Remove("AA");
        }

        if (options.Rectangle is not null) { annotation.Items["Rect"] = CreateNumberArray(options.Rectangle); invalidateAppearance = true; }
        if (options.QuadPoints is not null) { annotation.Items["QuadPoints"] = CreateNumberArray(options.QuadPoints); invalidateAppearance = true; }
        if (options.Vertices is not null) { annotation.Items["Vertices"] = CreateNumberArray(options.Vertices); invalidateAppearance = true; }
        if (options.Line is not null) { annotation.Items["L"] = CreateNumberArray(options.Line); invalidateAppearance = true; }
        if (options.InkPaths is not null) {
            var paths = new PdfArray(); foreach (IReadOnlyList<double> path in options.InkPaths) paths.Items.Add(CreateNumberArray(path)); annotation.Items["InkList"] = paths; invalidateAppearance = true;
        }
        if (options.LineStartEnding is not null || options.LineEndEnding is not null) {
            string start = options.LineStartEnding ?? ReadLineEnding(objects, annotation, 0) ?? "None";
            string end = options.LineEndEnding ?? ReadLineEnding(objects, annotation, 1) ?? "None";
            var endings = new PdfArray(); endings.Items.Add(new PdfName(start)); endings.Items.Add(new PdfName(end)); annotation.Items["LE"] = endings; invalidateAppearance = true;
        }
        if (options.InReplyToObjectNumber.HasValue) {
            int replyTarget = options.InReplyToObjectNumber.Value;
            if (!objects.TryGetValue(replyTarget, out PdfIndirectObject? target) || target.Value is not PdfDictionary targetDictionary || !IsAnnotation(targetDictionary)) throw new ArgumentException("Reply target annotation object was not found.", nameof(options));
            annotation.Items["IRT"] = new PdfReference(replyTarget, 0);
        }
        if (options.ReplyType is not null) annotation.Items["RT"] = new PdfName(options.ReplyType);
        if (options.PopupOpen.HasValue || options.PopupRectangle is not null) {
            PdfDictionary popup = ResolvePopup(objects, annotation, out int? popupObjectNumber);
            if (options.PopupOpen.HasValue) popup.Items["Open"] = new PdfBoolean(options.PopupOpen.Value);
            if (options.PopupRectangle is not null) popup.Items["Rect"] = CreateNumberArray(options.PopupRectangle);
            if (popupObjectNumber.HasValue) changedObjects.Add(popupObjectNumber.Value);
        }

        if (options.RegenerateAppearance) {
            changedObjects.Add(PdfAnnotationFlattener.RegenerateNormalAppearance(objects, annotation));
        } else if (invalidateAppearance) {
            annotation.Items.Remove("AP");
        }
        return changedObjects.AsReadOnly();
    }

    private static PdfArray CreateNumberArray(IReadOnlyList<double> values) { var array = new PdfArray(); foreach (double value in values) array.Items.Add(new PdfNumber(value)); return array; }
    private static string? ReadLineEnding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, int index) => annotation.Items.TryGetValue("LE", out PdfObject? value) && PdfObjectLookup.Resolve(objects, value) is PdfArray array && index < array.Items.Count && PdfObjectLookup.Resolve(objects, array.Items[index]) is PdfName name ? name.Name : null;
    private static PdfDictionary ResolvePopup(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, out int? objectNumber) {
        objectNumber = null;
        if (annotation.Get<PdfName>("Subtype")?.Name == "Popup") return annotation;
        if (annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) && PdfObjectLookup.Resolve(objects, popupObject) is PdfDictionary popup) { objectNumber = popupObject is PdfReference reference ? reference.ObjectNumber : null; return popup; }
        throw new InvalidOperationException("Annotation does not have a linked popup dictionary.");
    }

    private static PdfArray CreateColorArray(IReadOnlyList<double> values) {
        var color = new PdfArray();
        for (int i = 0; i < values.Count; i++) {
            color.Items.Add(new PdfNumber(ClampColor(values[i])));
        }

        return color;
    }

    private static double ClampColor(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }

    private static void ValidateRemovalOptions(PdfAnnotationRemovalOptions options) {
        if (options.ObjectNumber.HasValue && options.ObjectNumber.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Annotation object number must be positive.");
        }

        if (options.PageNumber.HasValue && options.PageNumber.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Page number must be positive.");
        }
    }

    private static void ValidateUpdateOptions(PdfAnnotationUpdateOptions options) {
        if (options.Color is not null) {
            if (options.Color.Count != 3) {
                throw new ArgumentException("Color must contain exactly three RGB values.", nameof(options));
            }

            for (int i = 0; i < options.Color.Count; i++) {
                double component = options.Color[i];
                if (double.IsNaN(component) || double.IsInfinity(component)) {
                    throw new ArgumentException("Color values must be finite RGB components.", nameof(options));
                }
            }
        }

        ValidateCoordinateArray(options.Rectangle, 4, 4, nameof(options.Rectangle));
        ValidateCoordinateArray(options.QuadPoints, 8, 0, nameof(options.QuadPoints));
        ValidateCoordinateArray(options.Vertices, 4, 0, nameof(options.Vertices));
        ValidateCoordinateArray(options.Line, 4, 4, nameof(options.Line));
        ValidateCoordinateArray(options.PopupRectangle, 4, 4, nameof(options.PopupRectangle));
        if (options.InkPaths is not null) { if (options.InkPaths.Count == 0) throw new ArgumentException("Ink paths cannot be empty.", nameof(options)); foreach (IReadOnlyList<double> path in options.InkPaths) ValidateCoordinateArray(path, 4, 0, nameof(options.InkPaths)); }
        if (options.InReplyToObjectNumber.HasValue && options.InReplyToObjectNumber.Value <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Reply parent object number must be positive.");
        if (options.ReplyType is not null && options.ReplyType != "R" && options.ReplyType != "Group") throw new ArgumentException("Reply type must be R or Group.", nameof(options));
        ValidatePdfName(options.LineStartEnding, nameof(options.LineStartEnding)); ValidatePdfName(options.LineEndEnding, nameof(options.LineEndEnding));

        if (options.Contents is null &&
            options.Title is null &&
            options.Name is null &&
            !options.Flags.HasValue &&
            options.Color is null &&
            !options.RemoveActions && options.Rectangle is null && options.QuadPoints is null && options.Vertices is null && options.Line is null && options.InkPaths is null && options.LineStartEnding is null && options.LineEndEnding is null && !options.InReplyToObjectNumber.HasValue && options.ReplyType is null && !options.PopupOpen.HasValue && options.PopupRectangle is null && !options.RegenerateAppearance) {
            throw new ArgumentException("At least one annotation update option must be provided.", nameof(options));
        }
    }

    private static void ValidateCoordinateArray(IReadOnlyList<double>? values, int minimum, int exact, string name) { if (values is null) return; if ((exact > 0 && values.Count != exact) || (exact == 0 && (values.Count < minimum || values.Count % 2 != 0))) throw new ArgumentException("Annotation coordinate array has an invalid length.", name); foreach (double value in values) if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Annotation coordinates must be finite.", name); }
    private static void ValidatePdfName(string? value, string name) { if (value is null) return; Guard.NotNullOrWhiteSpace(value, name); if (value.Any(char.IsWhiteSpace)) throw new ArgumentException("PDF name values cannot contain whitespace.", name); }

    private static bool IsAnnotation(PdfDictionary dictionary) {
        string? type = dictionary.Get<PdfName>("Type")?.Name;
        if (string.Equals(type, "Annot", StringComparison.Ordinal)) {
            return true;
        }

        if (!string.IsNullOrEmpty(type)) {
            return false;
        }

        return dictionary.Get<PdfName>("Subtype") is PdfName subtype &&
            KnownAnnotationSubtypes.Contains(subtype.Name);
    }

    private static bool HasAnnotationSubtype(PdfDictionary dictionary) {
        return dictionary.Items.ContainsKey("Subtype");
    }

    private static bool IsAnnotationUpdateTarget(Dictionary<int, PdfIndirectObject> objects, int objectNumber, PdfDictionary dictionary) {
        return IsAnnotation(dictionary) || IsReferencedFromPageAnnotations(objects, objectNumber);
    }

    private static bool IsReferencedFromPageAnnotations(Dictionary<int, PdfIndirectObject> objects, int objectNumber) {
        foreach (int pageObjectNumber in GetPageObjectNumbersInDocumentOrder(objects)) {
            if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary page ||
                !page.Items.TryGetValue("Annots", out PdfObject? annotsObject) ||
                PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annotations) {
                continue;
            }

            for (int i = 0; i < annotations.Items.Count; i++) {
                if (annotations.Items[i] is PdfReference reference &&
                    reference.ObjectNumber == objectNumber) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void RemovePopupReferences(Dictionary<int, PdfIndirectObject> objects, PdfArray annotations, HashSet<int> removed, ref bool changed) {
        if (removed.Count == 0) {
            return;
        }

        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            if (annotations.Items[i] is not PdfReference reference) {
                continue;
            }

            if (removed.Contains(reference.ObjectNumber)) {
                annotations.Items.RemoveAt(i);
                changed = true;
                continue;
            }

            if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfDictionary annotation &&
                annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
                popupObject is PdfReference popupReference &&
                removed.Contains(popupReference.ObjectNumber)) {
                annotation.Items.Remove("Popup");
                changed = true;
            }
        }
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            PdfObjectLookup.Resolve(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static List<int> GetPageObjectNumbersInDocumentOrder(Dictionary<int, PdfIndirectObject> objects) {
        var pages = new List<int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects);
        if (catalog is not null && catalog.Items.TryGetValue("Pages", out PdfObject? pagesRoot)) {
            CollectPageObjectNumbers(objects, pagesRoot, pages, new HashSet<int>());
        }

        if (pages.Count > 0) {
            return pages;
        }

        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.OrderBy(static item => item.Key)) {
            if (entry.Value.Value is PdfDictionary dictionary &&
                dictionary.Get<PdfName>("Type")?.Name == "Page") {
                pages.Add(entry.Key);
            }
        }

        return pages;
    }

    private static void CollectPageObjectNumbers(Dictionary<int, PdfIndirectObject> objects, PdfObject value, List<int> pages, HashSet<int> visitedObjects) {
        if (value is PdfReference reference) {
            if (!visitedObjects.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return;
            }

            if (indirect.Value is PdfDictionary referencedDictionary &&
                referencedDictionary.Get<PdfName>("Type")?.Name == "Page") {
                pages.Add(reference.ObjectNumber);
                return;
            }

            CollectPageObjectNumbers(objects, indirect.Value, pages, visitedObjects);
            return;
        }

        if (value is not PdfDictionary dictionary ||
            !dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        foreach (PdfObject kid in kids.Items) {
            CollectPageObjectNumbers(objects, kid, pages, visitedObjects);
        }
    }

    private static int FindCatalogObjectNumber(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
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

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf) {
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
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata))));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath = Path.GetFullPath(outputPath);
        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        return fullPath;
    }

    private static void WriteFile(string outputPath, byte[] bytes) {
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.WriteAllBytes(outputPath, bytes);
    }
}
