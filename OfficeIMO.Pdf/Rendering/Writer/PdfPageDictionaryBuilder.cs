using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfPageDictionaryBuilder {
    internal static string BuildGeneratedPageDictionary(
        int parentPagesId,
        double pageWidth,
        double pageHeight,
        int contentId,
        IReadOnlyList<(string Name, int Id)> fontResources,
        IReadOnlyList<(string Name, int Id)> xObjects,
        IReadOnlyList<(string Name, int Id)> graphicsStates,
        IReadOnlyList<(string Name, int Id)> shadings,
        IReadOnlyList<int> annotationIds,
        int? structParents = null,
        bool useStructureTabOrder = false,
        IReadOnlyList<(string Name, int Id)>? properties = null) {
        ValidatePositiveFinite(pageWidth, nameof(pageWidth));
        ValidatePositiveFinite(pageHeight, nameof(pageHeight));

        var sb = new StringBuilder();
        sb.Append("<< /Type /Page /Parent ")
            .Append(PdfSyntaxEscaper.IndirectReference(parentPagesId))
            .Append(" /MediaBox [0 0 ")
            .Append(FormatWholeNumber(pageWidth))
            .Append(' ')
            .Append(FormatWholeNumber(pageHeight))
            .Append("] /Resources <<");

        AppendResourcePart(sb, "Font", fontResources);
        AppendResourcePart(sb, "XObject", xObjects);
        AppendResourcePart(sb, "ExtGState", graphicsStates);
        AppendResourcePart(sb, "Shading", shadings);
        AppendResourcePart(sb, "Properties", properties ?? Array.Empty<(string Name, int Id)>());

        sb.Append(" >> /Contents ")
            .Append(PdfSyntaxEscaper.IndirectReference(contentId));

        AppendAnnotations(sb, annotationIds);
        if (structParents.HasValue) {
            sb.Append(" /StructParents ")
                .Append(structParents.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (useStructureTabOrder) {
            sb.Append(" /Tabs /S");
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string BuildResourcePart(string resourceKind, IReadOnlyList<(string Name, int Id)> resources) {
        var sb = new StringBuilder();
        AppendResourcePart(sb, resourceKind, resources);
        return sb.ToString();
    }

    internal static void AppendResourcePart(StringBuilder sb, string resourceKind, IReadOnlyList<(string Name, int Id)> resources) {
        Guard.NotNull(sb, nameof(sb));
        Guard.NotNullOrWhiteSpace(resourceKind, nameof(resourceKind));
        Guard.NotNull(resources, nameof(resources));
        if (resources.Count == 0) {
            return;
        }

        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(resourceKind))
            .Append(" << ");

        for (int i = 0; i < resources.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            AppendResourceName(sb, resources[i].Name);
            sb.Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(resources[i].Id));
        }

        sb.Append(" >>");
    }

    private static void AppendAnnotations(StringBuilder sb, IReadOnlyList<int> annotationIds) {
        Guard.NotNull(annotationIds, nameof(annotationIds));
        if (annotationIds.Count == 0) {
            return;
        }

        sb.Append(" /Annots [ ");
        for (int i = 0; i < annotationIds.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(PdfSyntaxEscaper.IndirectReference(annotationIds[i]));
        }

        sb.Append(" ]");
    }

    private static void AppendResourceName(StringBuilder sb, string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        string name = resourceName[0] == '/' ? resourceName.Substring(1) : resourceName;
        Guard.NotNullOrWhiteSpace(name, nameof(resourceName));
        sb.Append('/').Append(PdfSyntaxEscaper.Name(name));
    }

    private static string FormatWholeNumber(double value) =>
        value.ToString("0", CultureInfo.InvariantCulture);

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF page dimensions must be finite positive numbers.");
        }
    }
}
