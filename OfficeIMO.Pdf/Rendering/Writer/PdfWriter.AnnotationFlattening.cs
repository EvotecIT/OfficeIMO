namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFlattenedVisualAnnotationContent(
        LayoutResult.Page page,
        PdfOptions pageOptions,
        List<byte[]> objects,
        List<(string Name, int Id)> xobjects,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        Func<PdfOptions, int> ensureFormHelveticaFont,
        bool markAsArtifact) {
        Guard.NotNull(page, nameof(page));
        Guard.NotNull(pageOptions, nameof(pageOptions));
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(xobjects, nameof(xobjects));
        Guard.NotNull(ensureFont, nameof(ensureFont));
        Guard.NotNull(ensureFormHelveticaFont, nameof(ensureFormHelveticaFont));

        var sb = new StringBuilder();
        foreach (FreeTextAnnotation annotation in page.FreeTextAnnotations) {
            double width = annotation.X2 - annotation.X1;
            double height = annotation.Y2 - annotation.Y1;
            string appearanceContent = BuildFreeTextAnnotationAppearanceContent(
                annotation,
                width,
                height,
                pageOptions,
                ensureFont,
                out IReadOnlyList<(string Name, int Id)> appearanceFontResources);
            byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
            string appearanceDictionary = PdfAnnotationDictionaryBuilder.BuildAppearanceStreamDictionary(width, height, appearanceBytes.Length, appearanceFontResources);
            int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
            string resourceName = NextFlattenedAnnotationXObjectName(xobjects);
            xobjects.Add((resourceName, appearanceId));
            AppendFlattenedAnnotationDraw(sb, resourceName, annotation.X1, annotation.Y1);
        }

        foreach (HighlightAnnotation annotation in page.HighlightAnnotations) {
            double width = annotation.X2 - annotation.X1;
            double height = annotation.Y2 - annotation.Y1;
            string appearanceContent = PdfAnnotationDictionaryBuilder.BuildHighlightAppearanceContent(width, height, annotation.Color);
            byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
            string appearanceDictionary = PdfAnnotationDictionaryBuilder.BuildAppearanceStreamDictionary(width, height, appearanceBytes.Length, usesHighlightBlendMode: true);
            int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
            string resourceName = NextFlattenedAnnotationXObjectName(xobjects);
            xobjects.Add((resourceName, appearanceId));
            AppendFlattenedAnnotationDraw(sb, resourceName, annotation.X1, annotation.Y1);
        }

        return WrapArtifactContent(sb.ToString(), markAsArtifact);
    }

    private static void AppendFlattenedAnnotationDraw(StringBuilder sb, string resourceName, double x, double y) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .TransformMatrix(1D, 0D, 0D, 1D, x, y)
            .XObject(resourceName)
            .RestoreState();
    }

    private static string NextFlattenedAnnotationXObjectName(List<(string Name, int Id)> xobjects) {
        int index = 1;
        while (HasXObjectResourceName(xobjects, "/Ann" + index.ToString(System.Globalization.CultureInfo.InvariantCulture))) {
            index++;
        }

        return "/Ann" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static bool HasXObjectResourceName(List<(string Name, int Id)> xobjects, string name) {
        for (int i = 0; i < xobjects.Count; i++) {
            if (string.Equals(xobjects[i].Name, name, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
