namespace OfficeIMO.Pdf;

internal static class PdfStructTreeRootDictionaryBuilder {
    internal static string BuildEmptyStructTreeRootDictionary() {
        return "<< /Type /StructTreeRoot /K [] /RoleMap << >> >>\n";
    }

    internal static string BuildStructTreeRootDictionary(IReadOnlyList<int> childElementIds, int parentTreeId, int parentTreeNextKey) {
        Guard.NotNull(childElementIds, nameof(childElementIds));
        if (parentTreeNextKey < 0) {
            throw new ArgumentOutOfRangeException(nameof(parentTreeNextKey), parentTreeNextKey, "PDF parent-tree next key must be non-negative.");
        }

        var sb = new StringBuilder();
        sb.Append("<< /Type /StructTreeRoot /K ");
        AppendReferenceArray(sb, childElementIds);
        if (parentTreeId > 0) {
            sb.Append(" /ParentTree ")
                .Append(PdfSyntaxEscaper.IndirectReference(parentTreeId))
                .Append(" /ParentTreeNextKey ")
                .Append(parentTreeNextKey.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        sb.Append(" /RoleMap << >> >>\n");
        return sb.ToString();
    }

    internal static string BuildDocumentStructElement(int structTreeRootId, IReadOnlyList<int> childElementIds, string? language = null) {
        Guard.NotNull(childElementIds, nameof(childElementIds));
        var sb = new StringBuilder();
        sb.Append("<< /Type /StructElem /S /Document /P ")
            .Append(PdfSyntaxEscaper.IndirectReference(structTreeRootId))
            .Append(" /K ");
        AppendReferenceArray(sb, childElementIds);
        if (!string.IsNullOrWhiteSpace(language)) {
            sb.Append(" /Lang ")
                .Append(PdfSyntaxEscaper.TextString(language!));
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string BuildFigureStructElement(int parentId, int pageId, int markedContentId, string alternativeText) {
        Guard.NotNullOrWhiteSpace(alternativeText, nameof(alternativeText));
        return BuildStructElement(parentId, pageId, "Figure", markedContentId, alternativeText);
    }

    internal static string BuildTextStructElement(int parentId, int pageId, string structureType, int markedContentId, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
        Guard.NotNullOrWhiteSpace(structureType, nameof(structureType));
        return BuildStructElement(parentId, pageId, structureType, markedContentId, null, tableHeaderScope, tableColumnSpan, tableRowSpan);
    }

    internal static string BuildContainerStructElement(int parentId, int pageId, string structureType, IReadOnlyList<int> childElementIds) {
        Guard.NotNullOrWhiteSpace(structureType, nameof(structureType));
        Guard.NotNull(childElementIds, nameof(childElementIds));
        var sb = new StringBuilder();
        sb.Append("<< /Type /StructElem /S /")
            .Append(structureType)
            .Append(" /P ")
            .Append(PdfSyntaxEscaper.IndirectReference(parentId))
            .Append(" /Pg ")
            .Append(PdfSyntaxEscaper.IndirectReference(pageId))
            .Append(" /K ");
        AppendReferenceArray(sb, childElementIds);
        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string BuildAnnotationStructElement(int parentId, int pageId, int annotationObjectId, int? markedContentId = null, IReadOnlyList<int>? additionalMarkedContentIds = null, string structureType = "Link") {
        if (annotationObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(annotationObjectId), annotationObjectId, "PDF annotation object id must be positive.");
        }

        Guard.NotNullOrWhiteSpace(structureType, nameof(structureType));
        var sb = new StringBuilder();
        sb.Append("<< /Type /StructElem /S /")
            .Append(PdfSyntaxEscaper.Name(structureType))
            .Append(" /P ")
            .Append(PdfSyntaxEscaper.IndirectReference(parentId))
            .Append(" /Pg ")
            .Append(PdfSyntaxEscaper.IndirectReference(pageId))
            .Append(" /K ");
        if (markedContentId.HasValue) {
            sb.Append('[');
            AppendMarkedContentReference(sb, pageId, markedContentId.Value);
            if (additionalMarkedContentIds != null) {
                for (int i = 0; i < additionalMarkedContentIds.Count; i++) {
                    sb.Append(' ');
                    AppendMarkedContentReference(sb, pageId, additionalMarkedContentIds[i]);
                }
            }

            sb.Append(" << /Type /OBJR /Obj ")
                .Append(PdfSyntaxEscaper.IndirectReference(annotationObjectId))
                .Append(" >>]");
        } else {
            sb.Append("<< /Type /OBJR /Obj ")
                .Append(PdfSyntaxEscaper.IndirectReference(annotationObjectId))
                .Append(" >>");
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    private static void AppendMarkedContentReference(StringBuilder sb, int pageId, int markedContentId) {
        if (markedContentId < 0) {
            throw new ArgumentOutOfRangeException(nameof(markedContentId), markedContentId, "PDF marked-content id must be non-negative.");
        }

        sb.Append("<< /Type /MCR /Pg ")
            .Append(PdfSyntaxEscaper.IndirectReference(pageId))
            .Append(" /MCID ")
            .Append(markedContentId.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(" >>");
    }

    private static string BuildStructElement(int parentId, int pageId, string structureType, int markedContentId, string? alternativeText, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
        var sb = new StringBuilder();
        sb.Append("<< /Type /StructElem /S /")
            .Append(structureType)
            .Append(" /P ")
            .Append(PdfSyntaxEscaper.IndirectReference(parentId))
            .Append(" /Pg ")
            .Append(PdfSyntaxEscaper.IndirectReference(pageId))
            .Append(" /K << /Type /MCR /Pg ")
            .Append(PdfSyntaxEscaper.IndirectReference(pageId))
            .Append(" /MCID ")
            .Append(markedContentId.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(" >>");

        if (!string.IsNullOrWhiteSpace(alternativeText)) {
            sb.Append(" /Alt ")
                .Append(PdfSyntaxEscaper.TextString(alternativeText!));
        }

        if (ShouldEmitTableAttributes(structureType, tableHeaderScope, tableColumnSpan, tableRowSpan)) {
            AppendTableAttributes(sb, structureType, tableHeaderScope, tableColumnSpan, tableRowSpan);
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    private static bool ShouldEmitTableAttributes(string structureType, string tableHeaderScope, int tableColumnSpan, int tableRowSpan) {
        bool tableCell = string.Equals(structureType, "TH", StringComparison.Ordinal) || string.Equals(structureType, "TD", StringComparison.Ordinal);
        if (!tableCell) {
            return false;
        }

        return (string.Equals(structureType, "TH", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(tableHeaderScope)) ||
            tableColumnSpan > 1 ||
            tableRowSpan > 1;
    }

    private static void AppendTableAttributes(StringBuilder sb, string structureType, string tableHeaderScope, int tableColumnSpan, int tableRowSpan) {
        sb.Append(" /A << /O /Table");
        if (string.Equals(structureType, "TH", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(tableHeaderScope)) {
            sb.Append(" /Scope /")
                .Append(PdfSyntaxEscaper.Name(tableHeaderScope));
        }

        if (tableColumnSpan > 1) {
            sb.Append(" /ColSpan ")
                .Append(tableColumnSpan.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        if (tableRowSpan > 1) {
            sb.Append(" /RowSpan ")
                .Append(tableRowSpan.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        sb.Append(" >>");
    }

    internal static string BuildParentTree(IReadOnlyList<ParentTreeEntry> entries) {
        Guard.NotNull(entries, nameof(entries));
        var sb = new StringBuilder();
        sb.Append("<< /Nums [");
        for (int i = 0; i < entries.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(entries[i].StructParentIndex.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(' ');
            if (entries[i].IsArrayEntry) {
                AppendReferenceArray(sb, entries[i].StructElementIds);
            } else {
                sb.Append(PdfSyntaxEscaper.IndirectReference(entries[i].StructElementId));
            }
        }

        sb.Append("] >>\n");
        return sb.ToString();
    }

    private static void AppendReferenceArray(StringBuilder sb, IReadOnlyList<int> objectIds) {
        sb.Append('[');
        for (int i = 0; i < objectIds.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(PdfSyntaxEscaper.IndirectReference(objectIds[i]));
        }

        sb.Append(']');
    }

    internal sealed class ParentTreeEntry {
        private ParentTreeEntry(int structParentIndex, IReadOnlyList<int>? structElementIds, int structElementId) {
            if (structParentIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(structParentIndex), structParentIndex, "PDF parent-tree index must be non-negative.");
            }

            StructParentIndex = structParentIndex;
            StructElementIds = structElementIds ?? Array.Empty<int>();
            StructElementId = structElementId;
        }

        public int StructParentIndex { get; }

        public IReadOnlyList<int> StructElementIds { get; }

        public int StructElementId { get; }

        public bool IsArrayEntry => StructElementIds.Count > 0;

        public static ParentTreeEntry ForMarkedContentPage(int structParentIndex, IReadOnlyList<int> structElementIds) {
            Guard.NotNull(structElementIds, nameof(structElementIds));
            return new ParentTreeEntry(structParentIndex, structElementIds, 0);
        }

        public static ParentTreeEntry ForObjectReference(int structParentIndex, int structElementId) {
            if (structElementId <= 0) {
                throw new ArgumentOutOfRangeException(nameof(structElementId), structElementId, "PDF structure element id must be positive.");
            }

            return new ParentTreeEntry(structParentIndex, null, structElementId);
        }
    }
}
