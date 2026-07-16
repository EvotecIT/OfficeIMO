using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void ApplyMarkdownTextFallbackOptions(PdfCore.PdfOptions pdfOptions, MarkdownPdfSaveOptions options, MarkdownDoc document) {
        bool hasCallerPdfOptions = options.PdfOptions != null;
        bool hasExplicitFontFamily = !string.IsNullOrWhiteSpace(options.FontFamily);
        if (hasExplicitFontFamily) {
            pdfOptions.TryUseOfficeFontFamily(options.FontFamily, options.AllowSystemFontEmbedding);
        }

        if (options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) {
            return;
        }

        PdfCore.PdfTextFallbackFeatures fallbackFeatures = options.TextFallbacks;
        bool preserveDocumentFontSlots = hasCallerPdfOptions || hasExplicitFontFamily;
        bool usesCodeFont = MarkdownDocumentUsesCodeFont(document);
        if (!usesCodeFont) {
            fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.MonospaceFont;
        }

        if (preserveDocumentFontSlots) {
            fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.DocumentFont;
        }

        PdfCore.PdfTextFallbackFeatures documentAndMonospaceFallbacks =
            fallbackFeatures & (PdfCore.PdfTextFallbackFeatures.DocumentFont | PdfCore.PdfTextFallbackFeatures.MonospaceFont);
        if (documentAndMonospaceFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.UseTextFallbacks(
                documentAndMonospaceFallbacks,
                CreateMarkdownReservedFontSlots(pdfOptions, preserveDocumentFontSlots, reserveCourier: false),
                options.AllowSystemFontEmbedding);
        }

        PdfCore.PdfTextFallbackFeatures runFallbacks = fallbackFeatures &
            (PdfCore.PdfTextFallbackFeatures.MultilingualFonts | PdfCore.PdfTextFallbackFeatures.SymbolAndEmojiFonts);
        if (runFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.UseTextFallbacks(
                runFallbacks,
                CreateMarkdownReservedFontSlots(pdfOptions, preserveDocumentFontSlots, reserveCourier: usesCodeFont),
                options.AllowSystemFontEmbedding);
        }
    }

    private static bool MarkdownDocumentUsesCodeFont(MarkdownDoc document) {
        foreach (IMarkdownBlock block in document.DescendantsAndSelf()) {
            if (MarkdownBlockUsesCodeFont(block)) {
                return true;
            }
        }

        foreach (ListItem item in document.DescendantListItems()) {
            if (InlineSequenceUsesCodeFont(item.Content)) {
                return true;
            }

            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                if (InlineSequenceUsesCodeFont(item.AdditionalParagraphs[paragraphIndex])) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool MarkdownBlockUsesCodeFont(IMarkdownBlock block) {
        switch (block) {
            case CodeBlock:
            case SemanticFencedBlock:
                return true;
            case ParagraphBlock paragraph:
                return InlineSequenceUsesCodeFont(paragraph.Inlines);
            case HeadingBlock heading:
                return InlineSequenceUsesCodeFont(heading.Inlines);
            case CalloutBlock callout:
                return InlineSequenceUsesCodeFont(callout.TitleInlines);
            case DetailsBlock details:
                return details.Summary != null && InlineSequenceUsesCodeFont(details.Summary.Inlines);
            case DefinitionListBlock definitionList:
                return DefinitionListUsesCodeFont(definitionList);
            case TableBlock table:
                return TableUsesCodeFont(table);
            default:
                return false;
        }
    }

    private static bool DefinitionListUsesCodeFont(DefinitionListBlock definitionList) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        for (int i = 0; i < items.Count; i++) {
            if (InlineSequenceUsesCodeFont(items[i].Term) ||
                InlineSequenceUsesCodeFont(items[i].Definition)) {
                return true;
            }
        }

        return false;
    }

    private static bool TableUsesCodeFont(TableBlock table) {
        IReadOnlyList<InlineSequence> headers = table.HeaderInlines;
        for (int i = 0; i < headers.Count; i++) {
            if (InlineSequenceUsesCodeFont(headers[i])) {
                return true;
            }
        }

        IReadOnlyList<IReadOnlyList<InlineSequence>> rows = table.RowInlines;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IReadOnlyList<InlineSequence> row = rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                if (InlineSequenceUsesCodeFont(row[columnIndex])) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool InlineSequenceUsesCodeFont(InlineSequence sequence) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            if (InlineUsesCodeFont(sequence.Nodes[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool InlineUsesCodeFont(IMarkdownInline inline) {
        switch (inline) {
            case CodeSpanInline:
                return true;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                return InlineSequenceUsesCodeFont(container.NestedInlines);
            default:
                return false;
        }
    }

    private static IReadOnlyList<PdfCore.PdfStandardFont> CreateMarkdownReservedFontSlots(PdfCore.PdfOptions pdfOptions, bool includeDocumentFontSlots, bool reserveCourier) {
        var slots = new List<PdfCore.PdfStandardFont>();
        if (reserveCourier) {
            slots.Add(PdfCore.PdfStandardFont.Courier);
        }

        if (includeDocumentFontSlots) {
            slots.Add(pdfOptions.DefaultFont);
            slots.Add(pdfOptions.HeaderFont);
            slots.Add(pdfOptions.FooterFont);
        }

        return slots;
    }
}
