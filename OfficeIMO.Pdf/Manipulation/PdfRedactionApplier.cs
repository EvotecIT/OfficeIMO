using OfficeIMO.Drawing.Internal;
using OfficeIMO.Pdf.Filters;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Applies rectangle-based redactions by removing matched text objects and annotations, then painting redaction marks.
/// </summary>
internal static partial class PdfRedactionApplier {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    private static readonly Regex FontSelectionRegex = new Regex(@"/([^\s/]+)\s+[-+]?(?:\d+(?:\.\d+)?|\.\d+)\s+Tf\b", RegexOptions.Compiled, RegexTimeout);

    /// <summary>Applies a previously reviewed plan, including exact form-field removal for field-derived search areas.</summary>
    public static byte[] Apply(byte[] pdf, PdfRedactionPlan plan, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(plan, nameof(plan));
        if (plan.Areas.Count == 0) return (byte[])pdf.Clone();
        string[] fieldNames = plan.Areas.Select(static area => area.Label).Where(static label => label?.StartsWith("field:", StringComparison.Ordinal) == true).Select(static label => label!.Substring("field:".Length)).Distinct(StringComparer.Ordinal).ToArray();
        byte[] working = pdf;
        if (fieldNames.Length > 0) {
            var existing = new HashSet<string>(PdfReadDocument.Open(pdf, readOptions).FormFields.Where(static field => field.Name is not null).Select(static field => field.Name!), StringComparer.Ordinal);
            string[] removable = fieldNames.Where(existing.Contains).ToArray();
            if (removable.Length > 0) working = PdfAcroFormEditor.Edit(pdf, edit => { for (int i = 0; i < removable.Length; i++) edit.Remove(removable[i]); }, readOptions).ToBytes();
        }
        return Apply(working, plan.Areas, applyOptions, layoutOptions, readOptions);
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF byte array and returns rewritten PDF bytes.
    /// </summary>
    public static byte[] Apply(
        byte[] pdf,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(areas, nameof(areas));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.Redact, readOptions);

        PdfRedactionArea[] areaArray = areas.ToArray();
        if (areaArray.Length == 0) {
            throw new ArgumentException("At least one redaction area is required.", nameof(areas));
        }

        PdfRedactionApplyOptions effectiveOptions = applyOptions ?? new PdfRedactionApplyOptions();
        if (effectiveOptions.MaximumDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(applyOptions), "Maximum decoded image bytes must be positive.");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(pdf, areaArray, layoutOptions, readOptions);
        if (!plan.Preflight.CanReadLogicalObjects) {
            throw new InvalidOperationException("PDF redaction cannot be applied because logical content cannot be read. " + string.Join(" ", plan.Preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects)));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        ValidateRedactionAreas(areaArray, document.Pages.Count);
        int maximumDecodedStreamBytes = readOptions?.Limits.MaxDecodedStreamBytes ?? PdfReadLimits.DefaultMaxDecodedStreamBytes;
        RedactionMutation mutation = ApplyToObjects(objects, document, plan, areaArray, effectiveOptions, maximumDecodedStreamBytes);
        bool cleanupChanged = ApplyCleanupPolicy(objects, catalogObjectNumber, effectiveOptions.CleanupScope);
        if (!mutation.HasChanges && !cleanupChanged) {
            return pdf.ToArray();
        }

        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        PdfMetadata metadata = (effectiveOptions.CleanupScope & PdfRedactionCleanupScope.Metadata) != 0 ? new PdfMetadata() : document.UncheckedMetadata;
        return RewriteAllObjects(objects, catalogObjectNumber, metadata, pdf);
    }

    /// <summary>
    /// Applies rectangle-based redactions from the current position of a readable stream.
    /// </summary>
    public static byte[] Apply(
        Stream stream,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        return Apply(ReadStream(stream, nameof(stream)), areas, applyOptions, layoutOptions, readOptions);
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF and writes the rewritten bytes to a stream.
    /// </summary>
    public static void Apply(
        byte[] pdf,
        Stream outputStream,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        WriteOutput(outputStream, Apply(pdf, areas, applyOptions, layoutOptions, readOptions));
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF file and writes a new PDF file.
    /// </summary>
    public static void Apply(
        string inputPath,
        string outputPath,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] redacted = Apply(File.ReadAllBytes(inputPath), areas, applyOptions, layoutOptions, readOptions);
        WriteOutput(fullOutputPath, redacted);
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF file and returns rewritten PDF bytes.
    /// </summary>
    public static byte[] ApplyToBytes(
        string inputPath,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return Apply(File.ReadAllBytes(inputPath), areas, applyOptions, layoutOptions, readOptions);
    }

    private static RedactionMutation ApplyToObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadDocument document,
        PdfRedactionPlan plan,
        PdfRedactionArea[] areas,
        PdfRedactionApplyOptions options,
        int maximumDecodedStreamBytes) {
        var matchesByPage = plan.Matches
            .GroupBy(match => match.PageNumber)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var areasByPage = areas
            .GroupBy(area => area.PageNumber)
            .ToDictionary(group => group.Key, group => group.ToArray());
        bool changed = false;
        var removedImageObjectNumbers = new HashSet<int>();
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            int pageNumber = pageIndex + 1;
            PdfReadPage readPage = document.Pages[pageIndex];
            if (!objects.TryGetValue(readPage.ObjectNumber, out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            matchesByPage.TryGetValue(pageNumber, out PdfRedactionMatch[]? pageMatches);
            areasByPage.TryGetValue(pageNumber, out PdfRedactionArea[]? pageAreas);
            if ((pageMatches is null || pageMatches.Length == 0) &&
                (pageAreas is null || pageAreas.Length == 0)) {
                continue;
            }

            PdfRedactionMatch[] currentMatches = pageMatches ?? Array.Empty<PdfRedactionMatch>();
            ImageRedactionMutation imageMutation = RemoveMatchedImageObjects(objects, pageDictionary, currentMatches, options, ref nextObjectNumber);
            foreach (PdfRedactionMatch removedImage in imageMutation.RemovedMatches) if (removedImage.ObjectNumber.HasValue) removedImageObjectNumbers.Add(removedImage.ObjectNumber.Value);
            ValidateImagePlacementMatches(currentMatches, imageMutation.RemovedMatches, options);
            bool pageChanged = imageMutation.HasChanges;
            pageChanged = RemoveMatchedTextObjects(objects, pageDictionary, currentMatches, ref nextObjectNumber) || pageChanged;
            if (options.RemoveIntersectingPaths) pageChanged = RemoveIntersectingPathObjects(objects, pageDictionary, pageAreas ?? Array.Empty<PdfRedactionArea>(), maximumDecodedStreamBytes, ref nextObjectNumber) || pageChanged;
            pageChanged = RemoveMatchedAnnotations(objects, pageDictionary, currentMatches) || pageChanged;

            PdfRedactionArea[] paintAreas = SelectPaintAreas(pageAreas ?? Array.Empty<PdfRedactionArea>(), currentMatches, options);
            if (paintAreas.Length > 0) {
                IsolateExistingPageContents(objects, pageDictionary, ref nextObjectNumber);
                int contentObjectNumber = nextObjectNumber++;
                objects[contentObjectNumber] = new PdfIndirectObject(contentObjectNumber, 0, BuildRedactionContentStream(paintAreas, options.FillColor));
                AppendPageContent(objects, pageDictionary, contentObjectNumber);
                pageChanged = true;
            }

            changed = pageChanged || changed;
        }

        if (removedImageObjectNumbers.Count > 0) changed = RemoveUnusedImageObjectReferences(objects, removedImageObjectNumbers) || changed;

        return new RedactionMutation(changed);
    }

    private static void ValidateRedactionAreas(PdfRedactionArea[] areas, int pageCount) {
        for (int i = 0; i < areas.Length; i++) {
            if (areas[i].PageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(areas), "Redaction area page number " + areas[i].PageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page count " + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }
        }
    }

    private static void ValidateImagePlacementMatches(IReadOnlyList<PdfRedactionMatch> matches, IReadOnlyList<PdfRedactionMatch> removedMatches, PdfRedactionApplyOptions options) {
        if (options.AllowImagePlacementOverlays || options.UnsupportedImagePolicy == PdfRedactionUnsupportedImagePolicy.VisualOverlay) {
            return;
        }

        PdfRedactionMatch? imageMatch = matches.FirstOrDefault(match =>
            match.Kind == PdfRedactionMatchKind.ImagePlacement &&
            !ContainsRemovedImageMatch(removedMatches, match));
        if (imageMatch is null) {
            return;
        }

        string resourceName = string.IsNullOrEmpty(imageMatch.ResourceName) ? "unknown" : imageMatch.ResourceName!;
        throw new InvalidOperationException(
            "PDF redaction cannot be safely applied because a redaction area intersects image placement resource '" +
            resourceName +
            "' on page " +
            imageMatch.PageNumber.ToString(CultureInfo.InvariantCulture) +
            ". The image placement could not be rewritten safely; set PdfRedactionApplyOptions.AllowImagePlacementOverlays to true only when a visible overlay is an explicitly accepted weaker outcome.");
    }

    private static bool ContainsRemovedImageMatch(IReadOnlyList<PdfRedactionMatch> removedMatches, PdfRedactionMatch candidate) {
        if (removedMatches.Contains(candidate)) {
            return true;
        }

        for (int i = 0; i < removedMatches.Count; i++) {
            PdfRedactionMatch removed = removedMatches[i];
            if (removed.Kind == candidate.Kind &&
                removed.PageNumber == candidate.PageNumber &&
                string.Equals(removed.ResourceName, candidate.ResourceName, StringComparison.Ordinal) &&
                AreSameRedactionCoordinate(removed.X, candidate.X) &&
                AreSameRedactionCoordinate(removed.Y, candidate.Y) &&
                AreSameRedactionCoordinate(removed.Width, candidate.Width) &&
                AreSameRedactionCoordinate(removed.Height, candidate.Height)) {
                return true;
            }
        }

        return false;
    }

    private static bool AreSameRedactionCoordinate(double left, double right) =>
        Math.Abs(left - right) <= 0.001D;

    private static bool RemoveMatchedAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches) {
        PdfRedactionMatch[] annotationMatches = matches
            .Where(match => match.Kind == PdfRedactionMatchKind.Annotation)
            .ToArray();
        if (annotationMatches.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotsObject) ||
            PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annotations) {
            return false;
        }

        var annotationObjectNumbers = new HashSet<int>(annotationMatches
            .Where(match => match.ObjectNumber.HasValue)
            .Select(match => match.ObjectNumber!.Value));
        var popupObjectNumbers = new HashSet<int>();
        bool changed = false;
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            PdfReference? reference = item as PdfReference;
            PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
            if (annotation is null) {
                continue;
            }

            bool removeAnnotation =
                (reference is not null && annotationObjectNumbers.Contains(reference.ObjectNumber)) ||
                MatchesAnnotationRedaction(objects, annotation, annotationMatches);
            if (!removeAnnotation) {
                continue;
            }

            AddPopupObjectNumber(annotation, popupObjectNumbers);
            if (reference is not null) {
                annotationObjectNumbers.Add(reference.ObjectNumber);
                objects.Remove(reference.ObjectNumber);
            }

            annotations.Items.RemoveAt(i);
            changed = true;
        }

        RemovePopupReferences(objects, annotations, annotationObjectNumbers, popupObjectNumbers, ref changed);

        if (!changed) {
            return false;
        }

        if (annotations.Items.Count == 0) {
            pageDictionary.Items.Remove("Annots");
        }

        return true;
    }

    private static bool MatchesAnnotationRedaction(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        PdfRedactionMatch[] annotationMatches) {
        if (!TryReadRect(objects, annotation, out double x, out double y, out double width, out double height)) {
            return false;
        }

        string? subtype = TryReadName(objects, annotation, "Subtype");
        for (int i = 0; i < annotationMatches.Length; i++) {
            PdfRedactionMatch match = annotationMatches[i];
            if (!string.IsNullOrEmpty(match.Subtype) &&
                !string.Equals(match.Subtype, subtype, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (AreClose(match.X, x) &&
                AreClose(match.Y, y) &&
                AreClose(match.Width, width) &&
                AreClose(match.Height, height)) {
                return true;
            }
        }

        return false;
    }

    private static void AddPopupObjectNumber(PdfDictionary annotation, HashSet<int> popupObjectNumbers) {
        if (annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
            popupObject is PdfReference popupReference) {
            popupObjectNumbers.Add(popupReference.ObjectNumber);
        }
    }

    private static void RemovePopupReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray annotations,
        HashSet<int> removedAnnotationObjectNumbers,
        HashSet<int> popupObjectNumbers,
        ref bool changed) {
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            PdfReference? reference = item as PdfReference;
            PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
            if (annotation is null) {
                continue;
            }

            bool removePopup =
                (reference is not null && popupObjectNumbers.Contains(reference.ObjectNumber)) ||
                IsPopupForRemovedAnnotation(annotation, removedAnnotationObjectNumbers);
            if (!removePopup) {
                continue;
            }

            if (reference is not null) {
                objects.Remove(reference.ObjectNumber);
            }

            annotations.Items.RemoveAt(i);
            changed = true;
        }

        for (int i = 0; i < annotations.Items.Count; i++) {
            if (PdfObjectLookup.Resolve(objects, annotations.Items[i]) is not PdfDictionary annotation ||
                !annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) ||
                popupObject is not PdfReference popupReference ||
                (!removedAnnotationObjectNumbers.Contains(popupReference.ObjectNumber) &&
                !popupObjectNumbers.Contains(popupReference.ObjectNumber))) {
                continue;
            }

            annotation.Items.Remove("Popup");
            changed = true;
        }
    }

    private static bool IsPopupForRemovedAnnotation(PdfDictionary annotation, HashSet<int> removedAnnotationObjectNumbers) {
        if (!string.Equals(annotation.Get<PdfName>("Subtype")?.Name, "Popup", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return annotation.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
            parentObject is PdfReference parentReference &&
            removedAnnotationObjectNumbers.Contains(parentReference.ObjectNumber);
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            PdfObjectLookup.Resolve(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static bool TryReadRect(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        out double x,
        out double y,
        out double width,
        out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue("Rect", out PdfObject? rectObject) ||
            PdfObjectLookup.Resolve(objects, rectObject) is not PdfArray rect ||
            rect.Items.Count < 4 ||
            PdfObjectLookup.Resolve(objects, rect.Items[0]) is not PdfNumber x1 ||
            PdfObjectLookup.Resolve(objects, rect.Items[1]) is not PdfNumber y1 ||
            PdfObjectLookup.Resolve(objects, rect.Items[2]) is not PdfNumber x2 ||
            PdfObjectLookup.Resolve(objects, rect.Items[3]) is not PdfNumber y2) {
            return false;
        }

        double left = Math.Min(x1.Value, x2.Value);
        double right = Math.Max(x1.Value, x2.Value);
        double bottom = Math.Min(y1.Value, y2.Value);
        double top = Math.Max(y1.Value, y2.Value);
        if (double.IsNaN(left) || double.IsInfinity(left) ||
            double.IsNaN(right) || double.IsInfinity(right) ||
            double.IsNaN(bottom) || double.IsInfinity(bottom) ||
            double.IsNaN(top) || double.IsInfinity(top) ||
            right <= left ||
            top <= bottom) {
            return false;
        }

        x = left;
        y = bottom;
        width = right - left;
        height = top - bottom;
        return true;
    }

    private static PdfRedactionArea[] SelectPaintAreas(PdfRedactionArea[] areas, PdfRedactionMatch[] matches, PdfRedactionApplyOptions options) {
        if (options.PaintUnmatchedAreas) {
            return areas;
        }

        return areas
            .Where(area => matches.Any(match => ReferenceEquals(match.Area, area) || SameArea(match.Area, area)))
            .ToArray();
    }

    private static bool SameArea(PdfRedactionArea left, PdfRedactionArea right) {
        return left.PageNumber == right.PageNumber &&
            AreClose(left.X, right.X) &&
            AreClose(left.Y, right.Y) &&
            AreClose(left.Width, right.Width) &&
            AreClose(left.Height, right.Height);
    }

    private static bool AreClose(double left, double right) {
        return Math.Abs(left - right) < 0.0001D;
    }

    private static PdfStream BuildRedactionContentStream(PdfRedactionArea[] areas, PdfColor fillColor) {
        var builder = new StringBuilder();
        var content = new ContentStreamBuilder(builder)
            .SaveState()
            .FillColor(fillColor);
        for (int i = 0; i < areas.Length; i++) {
            content.Rectangle(areas[i].X, areas[i].Y, areas[i].Width, areas[i].Height)
                .FillPath();
        }

        content.RestoreState();
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(builder.ToString()));
    }

    private static void AppendPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, int contentObjectNumber) {
        var newReference = new PdfReference(contentObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            page.Items["Contents"] = newReference;
            return;
        }

        if (contents is PdfArray contentsArray) {
            contentsArray.Items.Add(newReference);
            return;
        }

        var array = new PdfArray();
        foreach (PdfObject item in EnumerateContentObjects(objects, contents)) {
            array.Items.Add(item);
        }

        array.Items.Add(newReference);
        page.Items["Contents"] = array;
    }

    private static IEnumerable<PdfReference> EnumerateContentReferences(Dictionary<int, PdfIndirectObject> objects, PdfObject contents) {
        foreach (PdfObject item in EnumerateContentObjects(objects, contents)) {
            if (item is PdfReference reference) {
                yield return reference;
            }
        }
    }

    private static IEnumerable<PdfObject> EnumerateContentObjects(Dictionary<int, PdfIndirectObject> objects, PdfObject contents) {
        if (contents is PdfArray directArray) {
            foreach (PdfObject item in directArray.Items) {
                yield return item;
            }

            yield break;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (PdfObject item in referencedArray.Items) {
                yield return item;
            }

            yield break;
        }

        yield return contents;
    }

    private static void ReplacePageContentReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary page,
        PdfObject contents,
        PdfReference source,
        PdfReference replacement) {
        if (ReferenceEquals(source, replacement)) {
            return;
        }

        if (contents is PdfReference reference && SameReference(reference, source)) {
            page.Items["Contents"] = replacement;
            return;
        }

        PdfArray? array = contents as PdfArray;
        if (array is null &&
            contents is PdfReference arrayReference &&
            PdfObjectLookup.TryGet(objects, arrayReference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfArray referencedArray) {
            array = CloneArray(referencedArray);
            page.Items["Contents"] = array;
        }

        if (array is null) {
            return;
        }

        for (int i = 0; i < array.Items.Count; i++) {
            if (array.Items[i] is PdfReference itemReference && SameReference(itemReference, source)) {
                array.Items[i] = replacement;
            }
        }
    }

    internal static void ReplacePageContentReferenceAtIndex(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary page,
        PdfObject contents,
        int contentIndex,
        PdfReference replacement) {
        if (contentIndex < 0) {
            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfArray referencedArray) {
            contents = CloneArray(referencedArray);
            page.Items["Contents"] = contents;
        }

        if (contents is PdfArray array) {
            int referenceIndex = 0;
            for (int itemIndex = 0; itemIndex < array.Items.Count; itemIndex++) {
                if (array.Items[itemIndex] is not PdfReference) {
                    continue;
                }

                if (referenceIndex == contentIndex) {
                    array.Items[itemIndex] = replacement;
                    return;
                }

                referenceIndex++;
            }

            return;
        }

        if (contentIndex == 0 && contents is PdfReference) {
            page.Items["Contents"] = replacement;
        }
    }

    private static void IsolateExistingPageContents(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, ref int nextObjectNumber) {
        if (!pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return;
        }

        PdfObject[] originalContents = EnumerateContentObjects(objects, contentsObject).ToArray();
        if (originalContents.Length == 0) {
            return;
        }

        int saveStateObjectNumber = nextObjectNumber++;
        int restoreStateObjectNumber = nextObjectNumber++;
        objects[saveStateObjectNumber] = new PdfIndirectObject(saveStateObjectNumber, 0, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("q\n")));
        objects[restoreStateObjectNumber] = new PdfIndirectObject(restoreStateObjectNumber, 0, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("\nQ\n")));

        var isolatedContents = new PdfArray();
        isolatedContents.Items.Add(new PdfReference(saveStateObjectNumber, 0));
        for (int i = 0; i < originalContents.Length; i++) {
            isolatedContents.Items.Add(originalContents[i]);
        }

        isolatedContents.Items.Add(new PdfReference(restoreStateObjectNumber, 0));
        pageDictionary.Items["Contents"] = isolatedContents;
    }

    private static PdfDictionary CleanStreamDictionary(PdfDictionary source) {
        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "Filter", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "DecodeParms", StringComparison.Ordinal)) {
                continue;
            }

            dictionary.Items[entry.Key] = entry.Value;
        }

        return dictionary;
    }

    private static Dictionary<int, int> CountIndirectReferenceUsage(Dictionary<int, PdfIndirectObject> objects) {
        var counts = new Dictionary<int, int>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            CountIndirectReferenceUsage(indirect.Value, counts, new HashSet<PdfObject>());
        }

        return counts;
    }

    private static void CountIndirectReferenceUsage(PdfObject value, Dictionary<int, int> counts, HashSet<PdfObject> visited) {
        if (!visited.Add(value)) {
            return;
        }

        switch (value) {
            case PdfReference reference:
                counts.TryGetValue(reference.ObjectNumber, out int count);
                counts[reference.ObjectNumber] = count + 1;
                break;
            case PdfArray array:
                foreach (PdfObject item in array.Items) {
                    CountIndirectReferenceUsage(item, counts, visited);
                }

                break;
            case PdfDictionary dictionary:
                foreach (PdfObject item in dictionary.Items.Values) {
                    CountIndirectReferenceUsage(item, counts, visited);
                }

                break;
            case PdfStream stream:
                CountIndirectReferenceUsage(stream.Dictionary, counts, visited);
                break;
        }
    }

    private static bool IsSharedReference(IReadOnlyDictionary<int, int> referenceCounts, PdfReference reference) =>
        referenceCounts.TryGetValue(reference.ObjectNumber, out int count) && count > 1;

    private static PdfReference CloneIndirectObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReference source,
        PdfIndirectObject indirect,
        ref int nextObjectNumber) {
        int objectNumber = nextObjectNumber++;
        var reference = new PdfReference(objectNumber, 0);
        objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, CloneObject(indirect.Value));
        return reference;
    }

    private static PdfObject CloneObject(PdfObject value) {
        switch (value) {
            case PdfReference reference:
                return new PdfReference(reference.ObjectNumber, reference.Generation);
            case PdfDictionary dictionary:
                return CloneDictionary(dictionary);
            case PdfArray array:
                return CloneArray(array);
            case PdfStream stream:
                return new PdfStream(CloneDictionary(stream.Dictionary), (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
            default:
                return value;
        }
    }

    private static PdfDictionary CloneDictionary(PdfDictionary source) {
        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
            dictionary.Items[entry.Key] = CloneObject(entry.Value);
        }

        return dictionary;
    }

    private static PdfArray CloneArray(PdfArray source) {
        var array = new PdfArray();
        foreach (PdfObject item in source.Items) {
            array.Items.Add(CloneObject(item));
        }

        return array;
    }

    private static bool SameReference(PdfReference left, PdfReference right) =>
        left.ObjectNumber == right.ObjectNumber && left.Generation == right.Generation;

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf) {
        int[] sourceIds = objects.Keys.OrderBy(id => id).ToArray();
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

    private static string NormalizeText(string value) {
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static bool ContainsOrdinal(string value, string match) {
#if NET472 || NETSTANDARD2_0
        return value.IndexOf(match, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(match, StringComparison.Ordinal);
#endif
    }

    private static string RemoveWhitespace(string value) {
        var builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) {
                builder.Append(value[i]);
            }
        }

        return builder.ToString();
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        string fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }

    private readonly struct RedactionMutation {
        public RedactionMutation(bool hasChanges) {
            HasChanges = hasChanges;
        }

        public bool HasChanges { get; }
    }

}
