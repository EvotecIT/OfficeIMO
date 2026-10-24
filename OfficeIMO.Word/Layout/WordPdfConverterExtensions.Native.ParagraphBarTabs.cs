using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool TryRenderNativeBarEmptyParagraph(INativePdfFlow flow, WordParagraph paragraph,
        PdfCore.PdfParagraphStyle style, NativeDocumentDefaults defaults, NativeFontMap fonts) {
        if (!style.TabStops.Any(stop => stop.Alignment == PdfCore.PdfTabAlignment.Bar)) return false;
        // A mark is a real line-box input even when it paints no glyph. This uses
        // the same font/spacing owner as empty table and note paragraphs.
        PdfCore.PdfTextRun mark = CreateNativeParagraphMarkStrut(paragraph,
            NativeTableStyleDefaults.Empty, defaults, fonts);
        flow.Paragraph(builder => builder.Runs(new[] { mark }), style: style);
        return true;
    }

    private static void RecordNativeParagraphBarTabDiagnostics(WordParagraph paragraph, WordToPdfOptions? options,
        string source, bool omitted = false, OpenXmlElement? identity = null) {
        if (options == null || IsNativeBarInMergeContinuation(paragraph) || !HasNativeParagraphBarTabs(paragraph) ||
            !ShouldRenderNativeEmptyParagraphLineBox(paragraph) &&
                !GetNativeRuns(paragraph).Any(run => IsNativeRenderableTextRun(run, paragraph)) ||
            !options.DiagnosedBarTabElements.Add(identity ?? paragraph._paragraph)) return;
        AddNativeExportWarning(options, omitted ? "NativeParagraphBarTabOmitted" : "NativeParagraphBarTabApproximated",
            source, omitted
                ? "Effective Word bar-tab rules cannot be retained by this PDF content projection."
                : "Word bar-tab rules retain their source position and per-line decoration; native vertical endpoints and device rounding are approximated.");
    }

    private static void RecordNativeRunningBarTabDiagnostics(WordSection section, WordToPdfOptions? options,
        bool mapsHeader, bool mapsFooter) {
        Record(section.ResolveDefaultHeader(), "default header", mapsHeader);
        Record(section.ResolveDefaultFooter(), "default footer", mapsFooter);
        if (section.DifferentFirstPage) {
            Record(section.ResolveFirstHeader(), "first header", mapsHeader);
            Record(section.ResolveFirstFooter(), "first footer", mapsFooter);
        }
        if (section.DocumentOddEvenSettingEnabled) {
            Record(section.ResolveEvenHeader(), "even header", mapsHeader);
            Record(section.ResolveEvenFooter(), "even footer", mapsFooter);
        }

        void Record(WordHeaderFooter? story, string source, bool mapped) {
            if (story == null || options == null) return;
            OpenXmlElement? root = (OpenXmlElement?)story._header ?? story._footer;
            if (root == null) return;
            if (!root.ChildElements.Any()) {
                // Empty referenced parts use the rich implicit-story owner even
                // when visible sibling stories use the caller's flat projection.
                RecordNativeParagraphBarTabDiagnostics(CreateNativeImplicitRunningParagraph(story),
                    options, source, identity: root);
                return;
            }
            foreach (W.Paragraph sourceParagraph in root.Descendants<W.Paragraph>()) {
                var paragraph = new WordParagraph(story.Document, sourceParagraph);
                if (!HasNativeParagraphBarTabs(paragraph)) continue;
                RecordNativeParagraphBarTabDiagnostics(paragraph,
                    options, source, omitted: !mapped || NativeBarParagraphProjectionOmitsRule(paragraph, section, options));
            }
        }
    }

    private static bool IsNativeBarInMergeContinuation(WordParagraph paragraph) =>
        paragraph._paragraph.Ancestors<W.TableCell>().Any(cell =>
            cell.TableCellProperties?.HorizontalMerge is { } horizontal && horizontal.Val?.Value != W.MergedCellValues.Restart ||
            cell.TableCellProperties?.VerticalMerge is { } vertical && vertical.Val?.Value != W.MergedCellValues.Restart);

    private static void RecordNativeCellBarTabDiagnostics(WordParagraph paragraph, WordToPdfOptions options, string source) {
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, source);
        // Table text keeps the cell's paragraph metadata, but flattens text-box
        // contents rather than projecting their inner paragraph decorations.
        foreach (W.Paragraph inner in paragraph._paragraph.Descendants<W.Paragraph>())
            RecordNativeParagraphBarTabDiagnostics(new WordParagraph(paragraph._document, inner),
                options, source + " text box", omitted: true);
    }

    private static bool NativeBarParagraphProjectionOmitsRule(WordParagraph paragraph,
        WordSection section, WordToPdfOptions? options) {
        if (paragraph._paragraph.Ancestors<W.TableCell>().Any())
            return paragraph._paragraph.Ancestors<W.TextBoxContent>().Any();
        if (WordParagraph.IsSectionMarkOnly(paragraph._paragraph)) return true;
        var runs = GetNativeRuns(paragraph);
        if (GetNativeVisibleTextBoxes(paragraph, runs).Count > 0)
            return !HasNativeMixedTextBoxOwnContent(paragraph, runs, section, options);
        if (IsNativeHorizontalRuleParagraph(paragraph, runs, paragraph.Text)) return true;
        // Object-only running paragraphs are emitted as flow objects. Unlike
        // body inline-picture paragraphs they have no separate rich text line.
        return !runs.Any(run => IsNativeRenderableTextRun(run, paragraph)) &&
            runs.Any(run => run.Image != null || run.Shape != null || run.Chart != null || run.TextBox != null);
    }

    private static void RenderNativeImplicitTerminalBodyBarParagraph(INativePdfFlow flow, WordSection section,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts) {
        // Word supplies a default paragraph mark after a terminal table even
        // though it is absent from the source. Only decorate it when a visible
        // effective bar requires paint; ordinary documents keep their prior flow.
        if (section._sectionProperties.Parent is not W.Body body ||
            body.ChildElements.LastOrDefault(element => element is not W.SectionProperties) is not W.Table) return;
        var paragraph = new WordParagraph(section._document, new W.Paragraph());
        if (!HasNativeParagraphBarTabs(paragraph) || !ShouldRenderNativeEmptyParagraphLineBox(paragraph)) return;
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, "implicit body paragraph after table", identity: body);
        RenderNativeEmptyParagraph(flow, paragraph, CreateNativeParagraphStyle(paragraph, defaults, fonts), defaults, fonts);
    }
}
