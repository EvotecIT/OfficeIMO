namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private int? RegisterStructureContainer(string structureType, int? parentElementIndex = null, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1, string? alternativeText = null) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int elementIndex = currentPage.StructElements.Count;
            currentPage.StructElements.Add(new PageStructElement {
                StructureType = structureType,
                ParentElementIndex = parentElementIndex,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                AlternativeText = alternativeText ?? string.Empty
            });
            return elementIndex;
        }

        private PageStructElement? RegisterStructureContainer(string structureType, PageStructElement? parentElement, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1, string? alternativeText = null) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            var element = new PageStructElement {
                StructureType = structureType,
                ParentElement = parentElement,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                AlternativeText = alternativeText ?? string.Empty
            };
            currentPage.StructElements.Add(element);
            return element;
        }

        private int? EnsurePageStructureContainer(string structureType, ref int? structureElementIndex, ref LayoutResult.Page? structurePage, int? parentElementIndex = null) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            if (!ReferenceEquals(structurePage, currentPage)) {
                structurePage = currentPage;
                structureElementIndex = RegisterStructureContainer(structureType, parentElementIndex);
            }

            return structureElementIndex;
        }

        private int? RegisterTextStructureElement(string structureType, int? parentElementIndex = null, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = structureType,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                ParentElementIndex = parentElementIndex
            });
            return markedContentId;
        }

        private int? RegisterTextStructureElement(string structureType, PageStructElement? parentElement, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = structureType,
                ParentElement = parentElement,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan
            });
            return markedContentId;
        }

        private int? RegisterFigureStructureElement(string alternativeText) {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = "Figure",
                AlternativeText = alternativeText
            });
            return markedContentId;
        }
    }
}
