using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

public static partial class DocumentReaderPdfExtensions {
    private static void AddTaggedContentMetadata(List<OfficeDocumentMetadataEntry> entries, PdfTaggedContentInfo? taggedContent) {
        if (taggedContent == null) {
            return;
        }

        AddCountMetadata(entries, "pdf-tagged-content-count", "pdf.taggedContent", "Count", 1);
        AddCountMetadata(entries, "pdf-tagged-content-structure-element-count", "pdf.taggedContent", "StructureElementCount", taggedContent.StructureElementCount);
        AddCountMetadata(entries, "pdf-tagged-content-parent-tree-entry-count", "pdf.taggedContent", "ParentTreeEntryCount", taggedContent.ParentTreeEntryCount);
        AddCountMetadata(entries, "pdf-tagged-content-marked-content-reference-count", "pdf.taggedContent", "MarkedContentReferenceCount", taggedContent.MarkedContentReferenceCount);
        AddCountMetadata(entries, "pdf-tagged-content-object-reference-count", "pdf.taggedContent", "ObjectReferenceCount", taggedContent.ObjectReferenceCount);
        AddCountMetadata(entries, "pdf-tagged-content-language-element-count", "pdf.taggedContent", "LanguageElementCount", taggedContent.LanguageElementCount);
        AddCountMetadata(entries, "pdf-tagged-content-alternate-text-element-count", "pdf.taggedContent", "AlternateTextElementCount", taggedContent.AlternateTextElementCount);
        AddCountMetadata(entries, "pdf-tagged-content-figure-without-alternate-text-count", "pdf.taggedContent", "FigureWithoutAlternateTextCount", taggedContent.FigureWithoutAlternateTextCount);

        foreach (KeyValuePair<string, int> count in taggedContent.StructureTypeCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            string typeKey = GetActionTypeKey(count.Key);
            AddCountMetadata(
                entries,
                "pdf-tagged-content-type-" + typeKey + "-count",
                "pdf.taggedContent",
                ToMetadataDisplayName(typeKey) + "StructureTypeCount",
                count.Value);
        }

        entries.Add(BuildTaggedContentMetadataEntry(taggedContent));
        for (int i = 0; i < taggedContent.StructureElements.Count; i++) {
            entries.Add(BuildTaggedContentElementMetadataEntry(taggedContent.StructureElements[i], i));
        }
    }

    private static OfficeDocumentMetadataEntry BuildTaggedContentMetadataEntry(PdfTaggedContentInfo taggedContent) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        AddAttribute(attributes, "marked", ToMetadataText(taggedContent.Marked));
        AddAttribute(attributes, "suspects", ToMetadataText(taggedContent.Suspects));
        AddAttribute(attributes, "userProperties", ToMetadataText(taggedContent.UserProperties));
        AddAttribute(attributes, "structTreeRootObjectNumber", taggedContent.StructTreeRootObjectNumber);
        AddAttribute(attributes, "parentTreeObjectNumber", taggedContent.ParentTreeObjectNumber);
        AddAttribute(attributes, "parentTreeNextKey", taggedContent.ParentTreeNextKey);
        AddAttribute(attributes, "rootElementObjectNumbers", FormatPdfIntegerComponents(taggedContent.RootElementObjectNumbers));
        AddAttribute(attributes, "parentTreeStructParentIndexes", FormatPdfIntegerComponents(taggedContent.ParentTreeStructParentIndexes));
        AddAttribute(attributes, "structureTypes", FormatPdfStringComponents(taggedContent.StructureTypes));
        AddAttribute(attributes, "hasRoleMap", ToMetadataText(taggedContent.HasRoleMap));
        AddAttribute(attributes, "hasDocumentStructureElement", ToMetadataText(taggedContent.HasDocumentStructureElement));
        AddAttribute(attributes, "hasMarkedContentReferences", ToMetadataText(taggedContent.HasMarkedContentReferences));
        AddAttribute(attributes, "hasObjectReferences", ToMetadataText(taggedContent.HasObjectReferences));
        AddAttribute(attributes, "hasDeepTaggedPdfEvidence", ToMetadataText(taggedContent.HasDeepTaggedPdfEvidence));
        AddAttribute(attributes, "figuresHaveAlternateText", ToMetadataText(taggedContent.FiguresHaveAlternateText));

        foreach (KeyValuePair<string, string> roleMapEntry in taggedContent.RoleMap.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddAttribute(attributes, "roleMap." + roleMapEntry.Key, roleMapEntry.Value);
        }

        return new OfficeDocumentMetadataEntry {
            Id = "pdf-tagged-content",
            Category = "pdf.taggedContent",
            Name = "TaggedContent",
            Value = taggedContent.Marked.HasValue ? ToMetadataText(taggedContent.Marked.Value) : null,
            ValueType = "object",
            SourceObjectId = taggedContent.StructTreeRootObjectNumber.HasValue
                ? taggedContent.StructTreeRootObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }

    private static OfficeDocumentMetadataEntry BuildTaggedContentElementMetadataEntry(PdfStructureElementInfo element, int elementIndex) {
        string id = "pdf-tagged-content-element-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["objectNumber"] = element.ObjectNumber.ToString(CultureInfo.InvariantCulture),
            ["markedContentReferenceCount"] = element.MarkedContentReferenceCount.ToString(CultureInfo.InvariantCulture),
            ["objectReferenceCount"] = element.ObjectReferenceCount.ToString(CultureInfo.InvariantCulture),
            ["hasChildElements"] = ToMetadataText(element.HasChildElements)
        };

        AddAttribute(attributes, "structureType", element.StructureType);
        AddAttribute(attributes, "parentObjectNumber", element.ParentObjectNumber);
        AddAttribute(attributes, "pageObjectNumber", element.PageObjectNumber);
        AddAttribute(attributes, "language", element.Language);
        AddAttribute(attributes, "alternateText", element.AlternateText);
        AddAttribute(attributes, "childElementObjectNumbers", FormatPdfIntegerComponents(element.ChildElementObjectNumbers));

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.taggedContent.element",
            Name = element.StructureType ?? "StructureElement",
            Value = element.Language ?? element.AlternateText,
            ValueType = "object",
            SourceObjectId = element.ObjectNumber.ToString(CultureInfo.InvariantCulture),
            Attributes = attributes
        };
    }
}
