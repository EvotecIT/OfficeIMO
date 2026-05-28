namespace OfficeIMO.Pdf;

internal static class PdfPageRangeObjectFilter {
    internal static int[] GetAllPageNumbers(int pageCount) {
        var pageNumbers = new int[pageCount];
        for (int i = 0; i < pageCount; i++) {
            pageNumbers[i] = i + 1;
        }

        return pageNumbers;
    }

    internal static bool ShouldUseDocumentWideObjects(int pageCount, int[] pageNumbers) {
        if (pageNumbers.Length != pageCount) {
            return false;
        }

        for (int i = 0; i < pageNumbers.Length; i++) {
            if (pageNumbers[i] != i + 1) {
                return false;
            }
        }

        return true;
    }

    internal static IReadOnlyList<PdfPageLabel> FilterPageLabelsByPageNumbers(IReadOnlyList<PdfPageLabel> pageLabels, int[] pageNumbers) {
        if (pageLabels.Count == 0) {
            return pageLabels;
        }

        var selectedSourceIndexes = new SortedSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            selectedSourceIndexes.Add(pageNumbers[i] - 1);
        }

        if (selectedSourceIndexes.Count == 0) {
            return Array.Empty<PdfPageLabel>();
        }

        var selectedLabels = new List<PdfPageLabel>();
        PdfPageLabel? previousSourceLabel = null;
        int previousSourceIndex = -1;
        foreach (int sourcePageIndex in selectedSourceIndexes) {
            PdfPageLabel? sourceLabel = FindPageLabelForSourceIndex(pageLabels, sourcePageIndex);
            if (sourceLabel is null) {
                continue;
            }

            bool continuesPreviousRun = previousSourceLabel is not null &&
                LabelsBelongToSameRule(previousSourceLabel, sourceLabel) &&
                sourcePageIndex == previousSourceIndex + 1;

            if (!continuesPreviousRun) {
                selectedLabels.Add(new PdfPageLabel(
                    sourcePageIndex,
                    sourceLabel.Style,
                    sourceLabel.Prefix,
                    GetAdjustedPageLabelStartNumber(sourceLabel, sourcePageIndex)));
            }

            previousSourceLabel = sourceLabel;
            previousSourceIndex = sourcePageIndex;
        }

        return selectedLabels.Count == 0 ? Array.Empty<PdfPageLabel>() : selectedLabels.AsReadOnly();
    }

    internal static IReadOnlyList<PdfOutlineItem> FilterOutlinesByPageNumbers(IReadOnlyList<PdfOutlineItem> outlines, int[] pageNumbers) {
        if (outlines.Count == 0) {
            return outlines;
        }

        var selectedPageNumbers = new HashSet<int>(pageNumbers);
        var selectedOutlines = new List<PdfOutlineItem>();
        for (int i = 0; i < outlines.Count; i++) {
            PdfOutlineItem? selected = FilterOutlineByPageNumbers(outlines[i], selectedPageNumbers);
            if (selected is not null) {
                selectedOutlines.Add(selected);
            }
        }

        return selectedOutlines.Count == 0 ? Array.Empty<PdfOutlineItem>() : selectedOutlines.AsReadOnly();
    }

    internal static IReadOnlyList<PdfNamedDestination> FilterNamedDestinationsByPageNumbers(IReadOnlyList<PdfNamedDestination> namedDestinations, int[] pageNumbers) {
        if (namedDestinations.Count == 0) {
            return namedDestinations;
        }

        var selectedPageNumbers = new HashSet<int>(pageNumbers);
        var selectedDestinations = new List<PdfNamedDestination>();
        for (int i = 0; i < namedDestinations.Count; i++) {
            PdfNamedDestination destination = namedDestinations[i];
            if (!destination.PageNumber.HasValue || selectedPageNumbers.Contains(destination.PageNumber.Value)) {
                selectedDestinations.Add(destination);
            }
        }

        return selectedDestinations.Count == 0 ? Array.Empty<PdfNamedDestination>() : selectedDestinations.AsReadOnly();
    }

    internal static PdfDocumentOpenAction? FilterOpenActionByPageNumbers(PdfDocumentOpenAction? openAction, int[] pageNumbers) {
        if (openAction is null || !openAction.PageNumber.HasValue) {
            return openAction;
        }

        var selectedPageNumbers = new HashSet<int>(pageNumbers);
        return selectedPageNumbers.Contains(openAction.PageNumber.Value) ? openAction : null;
    }

    internal static IReadOnlyList<PdfFormField> FilterFormFieldsByPageNumbers(IReadOnlyList<PdfFormField> formFields, int[] pageNumbers, bool preservePageDuplicates) {
        if (formFields.Count == 0) {
            return formFields;
        }

        var selectedPageNumbers = preservePageDuplicates ? null : new HashSet<int>(pageNumbers);
        var selectedFields = new List<PdfFormField>();

        for (int i = 0; i < formFields.Count; i++) {
            PdfFormField field = formFields[i];
            var selectedWidgets = new List<PdfFormWidget>();

            if (preservePageDuplicates) {
                for (int pageIndex = 0; pageIndex < pageNumbers.Length; pageIndex++) {
                    AddWidgetsForPage(field, pageNumbers[pageIndex], selectedWidgets);
                }
            } else {
                for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                    PdfFormWidget widget = field.Widgets[widgetIndex];
                    if (widget.PageNumber.HasValue && selectedPageNumbers!.Contains(widget.PageNumber.Value)) {
                        selectedWidgets.Add(widget);
                    }
                }
            }

            if (selectedWidgets.Count == 0) {
                continue;
            }

            selectedFields.Add(new PdfFormField(
                field.ObjectNumber,
                field.Name,
                field.PartialName,
                field.FieldType,
                field.Value,
                field.AlternateName,
                field.MappingName,
                field.Flags,
                field.MaxLength,
                field.Values,
                field.DefaultValue,
                field.DefaultValues,
                field.DefaultAppearance,
                field.Quadding,
                field.Options,
                selectedWidgets.AsReadOnly()));
        }

        return selectedFields.Count == 0 ? Array.Empty<PdfFormField>() : selectedFields.AsReadOnly();
    }

    private static void AddWidgetsForPage(PdfFormField field, int pageNumber, List<PdfFormWidget> selectedWidgets) {
        for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
            PdfFormWidget widget = field.Widgets[widgetIndex];
            if (widget.PageNumber == pageNumber) {
                selectedWidgets.Add(widget);
            }
        }
    }

    private static PdfPageLabel? FindPageLabelForSourceIndex(IReadOnlyList<PdfPageLabel> pageLabels, int sourcePageIndex) {
        PdfPageLabel? selected = null;
        for (int i = 0; i < pageLabels.Count; i++) {
            if (pageLabels[i].StartPageIndex > sourcePageIndex) {
                break;
            }

            selected = pageLabels[i];
        }

        return selected;
    }

    private static bool LabelsBelongToSameRule(PdfPageLabel left, PdfPageLabel right) {
        return left.StartPageIndex == right.StartPageIndex &&
            string.Equals(left.Style, right.Style, StringComparison.Ordinal) &&
            string.Equals(left.Prefix, right.Prefix, StringComparison.Ordinal) &&
            left.StartNumber == right.StartNumber;
    }

    private static int? GetAdjustedPageLabelStartNumber(PdfPageLabel sourceLabel, int sourcePageIndex) {
        if (sourceLabel.Style is null) {
            return sourceLabel.StartNumber;
        }

        int startNumber = sourceLabel.StartNumber ?? 1;
        return startNumber + sourcePageIndex - sourceLabel.StartPageIndex;
    }

    private static PdfOutlineItem? FilterOutlineByPageNumbers(PdfOutlineItem outline, HashSet<int> selectedPageNumbers) {
        var selectedChildren = new List<PdfOutlineItem>();
        for (int i = 0; i < outline.Children.Count; i++) {
            PdfOutlineItem? selectedChild = FilterOutlineByPageNumbers(outline.Children[i], selectedPageNumbers);
            if (selectedChild is not null) {
                selectedChildren.Add(selectedChild);
            }
        }

        bool keepOwnDestination = !outline.PageNumber.HasValue || selectedPageNumbers.Contains(outline.PageNumber.Value);
        if (!keepOwnDestination && selectedChildren.Count == 0) {
            return null;
        }

        return new PdfOutlineItem(
            outline.Title,
            outline.Level,
            keepOwnDestination ? outline.PageNumber : null,
            keepOwnDestination ? outline.DestinationTop : null,
            selectedChildren.Count == 0 ? Array.Empty<PdfOutlineItem>() : selectedChildren.AsReadOnly());
    }
}
