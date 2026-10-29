using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool TryRenderNativeSectionColumns(
        PdfCore.PdfPageBuilder page,
        WordSection section,
        IReadOnlyList<WordElement> elements,
        Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        Dictionary<WordParagraph, (int Level, int Index)> listIndices,
        NativeNoteNumbering footnoteNumbersById,
        WordPdfExportOptions? options,
        IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries,
        IReadOnlyDictionary<W.Paragraph, string> headingDestinations,
        NativeDocumentDefaults nativeDefaults,
        NativeFontMap nativeFontMap,
        bool balanceColumns) {
        int count = section.ColumnCount ?? 1;
        if (count <= 1) return false;
        double gap = section.ColumnsSpace is OfficeLength spacing && spacing >= OfficeLength.Zero
            ? spacing.ToPoints() : 36D;
        // Word 2013 layout ignores this legacy option without clearing its stored value.
        bool suppressBalancing = !UsesModernNativeWordLayout(section._document) &&
            section._document.CompatibilitySettings.DoNotBalanceTextColumns;
        var columnOptions = new PdfCore.PdfMultiColumnOptions {
            ColumnCount = count, Gap = gap,
            BalanceLastPage = balanceColumns && !suppressBalancing,
            BalanceKeptParagraphLines = !UsesModernNativeWordLayout(section._document),
            HonorKeepWithNextWhenBalancing = UsesModernNativeWordLayout(section._document),
            BalanceTableRowLines = UsesModernNativeWordLayout(section._document),
            FinalColumnSpacingAfter = GetNativeColumnSectionMarkHeight(section, elements, balanceColumns, nativeDefaults, nativeFontMap),
            SeparatorColor = section.HasColumnSeparator ? global::OfficeIMO.OfficeColor.Black : null,
            SeparatorWidth = section.HasColumnSeparator ? 0.5D : 0D
        };
        IReadOnlyList<WordSectionColumn> definitions = section.ColumnDefinitions;
        if (definitions.Count > 0) {
            columnOptions.ColumnDefinitions = definitions.Select(column => new PdfCore.PdfFlowColumn(
                PdfCore.PdfColumnWidth.Fixed(column.Width.ToPoints()),
                column.SpaceAfter?.ToPoints() ?? 0D)).ToArray();
        }
        PdfCore.PdfPageSize pageSize = GetNativePageSize(section, options);
        PdfCore.PdfPageMargins margins = GetNativeMargins(section, options);
        double sectionContentWidth = Math.Max(1D, pageSize.Width.ToPoints() - margins.Left.ToPoints() - margins.Right.ToPoints());
        double totalGap = definitions.Count > 0
            ? definitions.Take(count - 1).Sum(column => column.SpaceAfter?.ToPoints() ?? 0D)
            : gap * (count - 1);
        double firstColumnWidth = definitions.Count > 0
            ? definitions[0].Width.ToPoints()
            : (sectionContentWidth - totalGap) / count;
        page.Content(content => content.Columns(column => {
            INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(page, column, pageSize));
            for (int index = 0; index < elements.Count; index++) {
                WordElement element = elements[index];
                if (element is WordFootNote) continue;
                if (TryRenderNativeJoinedParagraphs(flow, elements, ref index,
                    paragraph => listMarkers.TryGetValue(paragraph, out var joinMarker) ? joinMarker : null,
                    footnoteNumbersById, options, nativeDefaults, nativeFontMap)) continue;
                if (TryRenderNativeList(flow, elements, ref index, listMarkers, listIndices, footnoteNumbersById, nativeDefaults, nativeFontMap, headingDestinations)) continue;
                RenderNativeElement(flow, element, section,
                    paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                    GetNativeFootnoteNumbersForElement(elements, index, footnoteNumbersById), footnoteNumbersById,
                    options, tableOfContentsEntries, headingDestinations, firstColumnWidth, nativeDefaults, nativeFontMap,
                    nextElement: GetNextNativeRenderableElement(elements, index));
            }
            RenderNativeImplicitTerminalBodyBarParagraph(flow, section, options, nativeDefaults, nativeFontMap);
        }, columnOptions));
        return true;
    }

    private static double GetNativeColumnSectionMarkHeight(WordSection section, IReadOnlyList<WordElement> elements,
        bool balanceColumns, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
        if (!balanceColumns || !UsesModernNativeWordLayout(section._document) || elements.Count < 2 ||
            elements[elements.Count - 1] is not WordParagraph paragraph ||
            elements[elements.Count - 2] is not WordTable || !WordParagraph.IsSectionMarkOnly(paragraph._paragraph)) return 0D;
        // Modern Word lays out the table's trailing section mark after balancing its row fragments.
        return MeasureNativeEmptyParagraphHeight(paragraph,
            CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap), nativeDefaults, nativeFontMap);
    }

}
