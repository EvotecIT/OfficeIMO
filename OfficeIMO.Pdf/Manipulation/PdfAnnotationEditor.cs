namespace OfficeIMO.Pdf;

/// <summary>Edits or removes PDF annotations without third-party dependencies.</summary>
public static class PdfAnnotationEditor {
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
    public static PdfAnnotationEditResult RemoveAnnotations(byte[] pdf, PdfAnnotationRemovalOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        PdfAnnotationRemovalOptions effectiveOptions = options ?? new PdfAnnotationRemovalOptions();
        ValidateRemovalOptions(effectiveOptions);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
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

        foreach (int objectNumber in removed) {
            objects.Remove(objectNumber);
        }

        if (!changed && removed.Count == 0) {
            return new PdfAnnotationEditResult((byte[])pdf.Clone(), 0);
        }

        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Load(pdf).Metadata, pdf);
        return new PdfAnnotationEditResult(rewritten, Math.Max(removed.Count, 1));
    }

    /// <summary>Updates a single indirect annotation and returns rewritten PDF bytes.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(byte[] pdf, int objectNumber, PdfAnnotationUpdateOptions options) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        if (objectNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "Annotation object number must be positive.");
        }

        ValidateUpdateOptions(options);
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfDictionary annotation ||
            !IsAnnotation(annotation)) {
            throw new ArgumentException("PDF annotation object was not found: " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(objectNumber));
        }

        ApplyUpdates(annotation, options);
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Load(pdf).Metadata, pdf);
        return new PdfAnnotationEditResult(rewritten, 1);
    }

    /// <summary>Removes annotations from a PDF file and writes the result to another file.</summary>
    public static PdfAnnotationEditResult RemoveAnnotations(string inputPath, string outputPath, PdfAnnotationRemovalOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        PdfAnnotationEditResult result = RemoveAnnotations(File.ReadAllBytes(inputPath), options);
        WriteFile(fullOutputPath, result.Bytes);
        return result;
    }

    /// <summary>Updates a single annotation in a PDF file and writes the result to another file.</summary>
    public static PdfAnnotationEditResult UpdateAnnotation(string inputPath, string outputPath, int objectNumber, PdfAnnotationUpdateOptions options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        PdfAnnotationEditResult result = UpdateAnnotation(File.ReadAllBytes(inputPath), objectNumber, options);
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

        return IsAnnotation(annotation);
    }

    private static void ApplyUpdates(PdfDictionary annotation, PdfAnnotationUpdateOptions options) {
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

        if (invalidateAppearance) {
            annotation.Items.Remove("AP");
        }
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
        if (options.Color is not null && options.Color.Count != 3) {
            throw new ArgumentException("Color must contain exactly three RGB values.", nameof(options));
        }

        if (options.Contents is null &&
            options.Title is null &&
            options.Name is null &&
            !options.Flags.HasValue &&
            options.Color is null &&
            !options.RemoveActions) {
            throw new ArgumentException("At least one annotation update option must be provided.", nameof(options));
        }
    }

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

        File.WriteAllBytes(outputPath, bytes);
    }
}
