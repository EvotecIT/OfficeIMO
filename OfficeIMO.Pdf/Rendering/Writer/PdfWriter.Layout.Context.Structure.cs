namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private int? RegisterStructureContainer(string structureType, int? parentElementIndex = null, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1, string? alternativeText = null) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            PageStructElement? semanticParent = parentElementIndex.HasValue ? null : ResolveFlowSemanticParent();
            int elementIndex = currentPage.StructElements.Count;
            currentPage.StructElements.Add(new PageStructElement {
                StructureType = structureType,
                ParentElementIndex = parentElementIndex,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                AlternativeText = alternativeText ?? string.Empty,
                ParentElement = semanticParent
            });
            return elementIndex;
        }

        private PageStructElement? RegisterStructureContainer(string structureType, PageStructElement? parentElement, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1, string? alternativeText = null) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            var element = new PageStructElement {
                StructureType = structureType,
                ParentElement = parentElement ?? ResolveFlowSemanticParent(),
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                AlternativeText = alternativeText ?? string.Empty
            };
            currentPage.StructElements.Add(element);
            return element;
        }

        private int? EnsurePageStructureContainer(string structureType, ref int? structureElementIndex, ref LayoutResult.Page? structurePage, int? parentElementIndex = null) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            if (!ReferenceEquals(structurePage, currentPage)) {
                structurePage = currentPage;
                structureElementIndex = RegisterStructureContainer(structureType, parentElementIndex);
            }

            return structureElementIndex;
        }

        private int? RegisterTextStructureElement(string structureType, int? parentElementIndex = null, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            PageStructElement? semanticParent = parentElementIndex.HasValue ? null : ResolveFlowSemanticParent();
            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = structureType,
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan,
                ParentElementIndex = parentElementIndex,
                ParentElement = semanticParent
            });
            return markedContentId;
        }

        private int? RegisterTextStructureElement(string structureType, PageStructElement? parentElement, string tableHeaderScope = "", int tableColumnSpan = 1, int tableRowSpan = 1) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = structureType,
                ParentElement = parentElement ?? ResolveFlowSemanticParent(),
                TableHeaderScope = tableHeaderScope,
                TableColumnSpan = tableColumnSpan,
                TableRowSpan = tableRowSpan
            });
            return markedContentId;
        }

        private int? RegisterFigureStructureElement(string alternativeText, int? parentElementIndex = null) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            PageStructElement? semanticParent = parentElementIndex.HasValue ? null : ResolveFlowSemanticParent();
            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = "Figure",
                AlternativeText = alternativeText,
                ParentElementIndex = parentElementIndex,
                ParentElement = semanticParent
            });
            return markedContentId;
        }

        private int? RegisterFigureStructureElement(string alternativeText, PageStructElement? parentElement) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null) {
                return null;
            }

            int markedContentId = currentPage.NextMarkedContentId++;
            currentPage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = "Figure",
                AlternativeText = alternativeText,
                ParentElement = parentElement ?? ResolveFlowSemanticParent()
            });
            return markedContentId;
        }

        private PageStructElement? ResolveFlowSemanticParent() {
            if (!emitGeneratedStructure || currentPage == null || flowSemanticScopes.Count == 0) {
                return null;
            }

            PageStructElement? parent = null;
            for (int index = 0; index < flowSemanticScopes.Count; index++) {
                FlowSemanticScope scope = flowSemanticScopes[index];
                PageStructElement? element = scope.Element;
                if (element == null) {
                    element = new PageStructElement {
                        StructureType = MapSemanticStructureType(scope.Role),
                        AlternativeText = scope.AlternativeText ?? string.Empty,
                        ParentElement = parent
                    };
                    currentPage.StructElements.Add(element);
                    scope.Element = element;
                    scope.ElementPage = currentPage;
                } else if (!ReferenceEquals(scope.ElementPage, currentPage)) {
                    element.SpansPages = true;
                }

                parent = element;
            }

            return parent;
        }
    }
}
