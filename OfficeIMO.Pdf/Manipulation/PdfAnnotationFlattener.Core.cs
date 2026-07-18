namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    private static int FlattenPageVisualAnnotations(Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber, PdfAnnotationFlattenOptions? options, Dictionary<int, int> pageNumbers) {
        int flattenedCount = 0;
        foreach (var entry in objects.OrderBy(pair => pair.Key).ToArray()) {
            if (entry.Value.Value is not PdfDictionary page ||
                page.Get<PdfName>("Type")?.Name != "Page" ||
                !page.Items.TryGetValue("Annots", out var annotsObject) ||
                ResolveObject(objects, annotsObject) is not PdfArray annots) {
                continue;
            }
            if (options?.PageNumber != null && (!pageNumbers.TryGetValue(entry.Key, out int pageNumber) || pageNumber != options.PageNumber.Value)) continue;

            var pageAnnotations = new List<FlattenVisualAnnotationState>();
            var remainingAnnots = new PdfArray();
            var flattenedAnnotationObjectNumbers = new HashSet<int>();
            for (int i = 0; i < annots.Items.Count; i++) {
                PdfObject annotObject = annots.Items[i];
                PdfDictionary? annotation = ResolveDictionary(objects, annotObject);
                if (annotation == null || TryReadName(objects, annotation, "Subtype") is not string subtype ||
                    !MatchesFlattenSelector(annotObject, subtype, options) ||
                    !IsSupportedVisualAnnotation(subtype)) {
                    remainingAnnots.Items.Add(annotObject);
                    continue;
                }

                if (HasNonViewableAnnotationFlag(objects, annotation)) {
                    remainingAnnots.Items.Add(annotObject);
                    continue;
                }

                if (!TryReadRectCoordinates(objects, annotation, out double x, out double y, out double width, out double height)) {
                    throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
                }

                if (!TryGetNormalAppearanceReference(objects, annotation, out PdfReference? appearanceReference)) {
                    appearanceReference = CreateSyntheticAppearanceReference(objects, annotation, subtype, x, y, width, height, ref nextObjectNumber);
                }

                AppearancePlacement placement = ReadAppearancePlacement(objects, appearanceReference!, x, y, width, height);
                pageAnnotations.Add(new FlattenVisualAnnotationState(placement, appearanceReference!.ObjectNumber));
                if (annotObject is PdfReference annotationReference) {
                    flattenedAnnotationObjectNumbers.Add(annotationReference.ObjectNumber);
                    objects.Remove(annotationReference.ObjectNumber);
                }
            }

            if (pageAnnotations.Count == 0) {
                continue;
            }

            RemovePopupAnnotationsForFlattenedParents(objects, remainingAnnots, flattenedAnnotationObjectNumbers);

            if (remainingAnnots.Items.Count == 0) {
                page.Items.Remove("Annots");
            } else {
                page.Items["Annots"] = remainingAnnots;
            }

            string content = BuildFlattenContent(objects, page, pageAnnotations);
            int contentObjectNumber = nextObjectNumber++;
            objects[contentObjectNumber] = new PdfIndirectObject(contentObjectNumber, 0, CreateContentStream(content));
            AppendPageContent(objects, page, contentObjectNumber);
            flattenedCount += pageAnnotations.Count;
        }

        return flattenedCount;
    }

    private static bool MatchesFlattenSelector(PdfObject annotation, string subtype, PdfAnnotationFlattenOptions? options) {
        if (options == null) return true;
        if (options.ObjectNumber.HasValue && (annotation is not PdfReference reference || reference.ObjectNumber != options.ObjectNumber.Value)) return false;
        return options.Subtype == null || string.Equals(options.Subtype, subtype, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportedVisualAnnotation(string subtype) {
        return string.Equals(subtype, "FreeText", StringComparison.Ordinal) ||
            string.Equals(subtype, "Highlight", StringComparison.Ordinal) ||
            string.Equals(subtype, "Underline", StringComparison.Ordinal) ||
            string.Equals(subtype, "StrikeOut", StringComparison.Ordinal) ||
            string.Equals(subtype, "Squiggly", StringComparison.Ordinal) ||
            string.Equals(subtype, "Square", StringComparison.Ordinal) ||
            string.Equals(subtype, "Circle", StringComparison.Ordinal) ||
            string.Equals(subtype, "Line", StringComparison.Ordinal) ||
            string.Equals(subtype, "Ink", StringComparison.Ordinal) ||
            string.Equals(subtype, "Polygon", StringComparison.Ordinal) ||
            string.Equals(subtype, "PolyLine", StringComparison.Ordinal) ||
            string.Equals(subtype, "Stamp", StringComparison.Ordinal) ||
            string.Equals(subtype, "Caret", StringComparison.Ordinal);
    }

    private static bool HasNonViewableAnnotationFlag(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        const int invisible = 1;
        const int hidden = 2;
        const int noView = 32;

        if (!annotation.Items.TryGetValue("F", out var flagsObject) ||
            ResolveObject(objects, flagsObject) is not PdfNumber flagsNumber) {
            return false;
        }

        int flags = (int)flagsNumber.Value;
        return (flags & (invisible | hidden | noView)) != 0;
    }

    private static string BuildFlattenContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, List<FlattenVisualAnnotationState> annotations) {
        PdfDictionary xObjects = EnsurePageXObjects(objects, page);
        var builder = new StringBuilder();
        for (int i = 0; i < annotations.Count; i++) {
            FlattenVisualAnnotationState annotation = annotations[i];
            string xObjectName = CreateUniqueXObjectName(xObjects);
            xObjects.Items[xObjectName] = new PdfReference(annotation.AppearanceObjectNumber, 0);
            builder.Append("q\n");
            builder.Append(FormatNumber(annotation.Placement.A))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.B))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.C))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.D))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.E))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.F))
                .Append(" cm\n");
            builder.Append('/').Append(xObjectName).Append(" Do\n");
            builder.Append("Q\n");
        }

        return builder.ToString();
    }

    private static void RemovePopupAnnotationsForFlattenedParents(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray remainingAnnots,
        HashSet<int> flattenedAnnotationObjectNumbers) {
        if (flattenedAnnotationObjectNumbers.Count == 0 || remainingAnnots.Items.Count == 0) {
            return;
        }

        for (int i = remainingAnnots.Items.Count - 1; i >= 0; i--) {
            PdfObject popupObject = remainingAnnots.Items[i];
            PdfDictionary? popup = ResolveDictionary(objects, popupObject);
            if (popup is null ||
                !string.Equals(TryReadName(objects, popup, "Subtype"), "Popup", StringComparison.Ordinal) ||
                !popup.Items.TryGetValue("Parent", out var parentObject) ||
                parentObject is not PdfReference parentReference ||
                !flattenedAnnotationObjectNumbers.Contains(parentReference.ObjectNumber)) {
                continue;
            }

            remainingAnnots.Items.RemoveAt(i);
            if (popupObject is PdfReference popupReference) {
                objects.Remove(popupReference.ObjectNumber);
            }
        }
    }

    private sealed class FlattenVisualAnnotationState {
        public FlattenVisualAnnotationState(AppearancePlacement placement, int appearanceObjectNumber) {
            Placement = placement;
            AppearanceObjectNumber = appearanceObjectNumber;
        }

        public AppearancePlacement Placement { get; }
        public int AppearanceObjectNumber { get; }
    }

    private sealed class AppearancePlacement {
        public AppearancePlacement(double a, double b, double c, double d, double e, double f) {
            A = a;
            B = b;
            C = c;
            D = d;
            E = e;
            F = f;
        }

        public double A { get; }
        public double B { get; }
        public double C { get; }
        public double D { get; }
        public double E { get; }
        public double F { get; }
    }
}
