namespace OfficeIMO.Pdf;

public static partial class PdfAnnotationFlattener {
    private static int FlattenPageVisualAnnotations(Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber) {
        int flattenedCount = 0;
        foreach (var entry in objects.OrderBy(pair => pair.Key).ToArray()) {
            if (entry.Value.Value is not PdfDictionary page ||
                page.Get<PdfName>("Type")?.Name != "Page" ||
                !page.Items.TryGetValue("Annots", out var annotsObject) ||
                ResolveObject(objects, annotsObject) is not PdfArray annots) {
                continue;
            }

            var pageAnnotations = new List<FlattenVisualAnnotationState>();
            var remainingAnnots = new PdfArray();
            for (int i = 0; i < annots.Items.Count; i++) {
                PdfObject annotObject = annots.Items[i];
                PdfDictionary? annotation = ResolveDictionary(objects, annotObject);
                if (annotation == null ||
                    TryReadName(objects, annotation, "Subtype") is not string subtype ||
                    !IsSupportedVisualAnnotation(subtype)) {
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
                    objects.Remove(annotationReference.ObjectNumber);
                }
            }

            if (pageAnnotations.Count == 0) {
                continue;
            }

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

    private static string BuildFlattenContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, List<FlattenVisualAnnotationState> annotations) {
        PdfDictionary xObjects = EnsurePageXObjects(objects, page);
        var builder = new StringBuilder();
        for (int i = 0; i < annotations.Count; i++) {
            FlattenVisualAnnotationState annotation = annotations[i];
            string xObjectName = CreateUniqueXObjectName(xObjects);
            xObjects.Items[xObjectName] = new PdfReference(annotation.AppearanceObjectNumber, 0);
            builder.Append("q\n");
            builder.Append(FormatNumber(annotation.Placement.ScaleX))
                .Append(" 0 0 ")
                .Append(FormatNumber(annotation.Placement.ScaleY))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.TranslateX))
                .Append(' ')
                .Append(FormatNumber(annotation.Placement.TranslateY))
                .Append(" cm\n");
            builder.Append('/').Append(xObjectName).Append(" Do\n");
            builder.Append("Q\n");
        }

        return builder.ToString();
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
        public AppearancePlacement(double scaleX, double scaleY, double translateX, double translateY) {
            ScaleX = scaleX;
            ScaleY = scaleY;
            TranslateX = translateX;
            TranslateY = translateY;
        }

        public double ScaleX { get; }
        public double ScaleY { get; }
        public double TranslateX { get; }
        public double TranslateY { get; }
    }
}
