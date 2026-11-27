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
        flow.Paragraph(builder => builder.AddRuns(new[] { mark }), style: style);
        return true;
    }

    private static void RecordNativeParagraphBarTabDiagnostics(WordParagraph paragraph, WordPdfExportOptions? options,
        string source, bool omitted = false, OfficeXmlNode identity = default) {
        if (options == null || IsNativeBarInMergeContinuation(paragraph) || !HasNativeParagraphBarTabs(paragraph) ||
            !ShouldRenderNativeEmptyParagraphLineBox(paragraph) &&
                !GetNativeRuns(paragraph).Any(run => IsNativeRenderableTextRun(run, paragraph)) ||
            !options.DiagnosedBarTabElements.Add(identity.Exists ? identity : paragraph.Node())) return;
        AddNativeExportWarning(options, omitted ? "native-paragraph-bar-tab-omitted" : "native-paragraph-bar-tab-approximated",
            source, omitted
                ? "Effective Word bar-tab rules cannot be retained by this PDF content projection."
                : "Word bar-tab rules retain their source position and per-line decoration; native vertical endpoints and device rounding are approximated.");
    }

    private static void RecordNativeRunningBarTabDiagnostics(WordSection section, WordPdfExportOptions? options,
        bool mapsHeader, bool mapsFooter) {
        Record(section.ResolveDefaultHeader(), "default header", mapsHeader);
        Record(section.ResolveDefaultFooter(), "default footer", mapsFooter);
        if (section.PageSetup.DifferentFirstPage) {
            Record(section.ResolveFirstHeader(), "first header", mapsHeader);
            Record(section.ResolveFirstFooter(), "first footer", mapsFooter);
        }
        if (section.DocumentOddEvenSettingEnabled) {
            Record(section.ResolveEvenHeader(), "even header", mapsHeader);
            Record(section.ResolveEvenFooter(), "even footer", mapsFooter);
        }

        void Record(WordHeaderFooter? story, string source, bool mapped) {
            if (story == null || options == null) return;
            OfficeXmlNode root = story.Node();
            if (root.IsNone) return;
            if (!root.HasChildren()) {
                // Empty referenced parts use the rich implicit-story owner even
                // when visible sibling stories use the caller's flat projection.
                RecordNativeParagraphBarTabDiagnostics(CreateNativeImplicitRunningParagraph(story),
                    options, source, identity: root);
                return;
            }
            foreach (OfficeXmlNode sourceParagraph in root.ModelDescendants("p")) {
                WordParagraph paragraph = story.Document.ParagraphView(sourceParagraph);
                if (!HasNativeParagraphBarTabs(paragraph)) continue;
                RecordNativeParagraphBarTabDiagnostics(paragraph,
                    options, source, omitted: !mapped || NativeBarParagraphProjectionOmitsRule(paragraph, section, options));
            }
        }
    }

    private static bool IsNativeBarInMergeContinuation(WordParagraph paragraph) =>
        paragraph.Node().Ancestors().Where(cell => cell.IsW("tc")).Any(cell =>
            cell.WChild("tcPr").WChild("hMerge") is { Exists: true } horizontal && horizontal.WVal() != "restart" ||
            cell.WChild("tcPr").WChild("vMerge") is { Exists: true } vertical && vertical.WVal() != "restart");

    private static void RecordNativeCellBarTabDiagnostics(WordParagraph paragraph, WordPdfExportOptions options, string source) {
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, source);
        // Table text keeps the cell's paragraph metadata, but flattens text-box
        // contents rather than projecting their inner paragraph decorations.
        foreach (OfficeXmlNode inner in paragraph.Node().ModelDescendants("p"))
            RecordNativeParagraphBarTabDiagnostics(paragraph._document.ParagraphView(inner),
                options, source + " text box", omitted: true);
    }

    private static bool NativeBarParagraphProjectionOmitsRule(WordParagraph paragraph,
        WordSection section, WordPdfExportOptions? options) {
        if (paragraph.Node().Ancestor(WordXml.W, "tc").Exists)
            return paragraph.Node().Ancestor(WordXml.W, "txbxContent").Exists;
        if (IsSectionMarkOnly(paragraph.Node())) return true;
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
        WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts) {
        // Word supplies a default paragraph mark after a terminal table even
        // though it is absent from the source. Only decorate it when a visible
        // effective bar requires paint; ordinary documents keep their prior flow.
        OfficeXmlNode body = section.Node().Parent;
        if (!body.IsW("body") || !body.Elements().LastOrDefault(element => !element.IsW("sectPr")).IsW("tbl")) return;
        WordParagraph paragraph = section._document.CreateDetachedParagraphView();
        if (!HasNativeParagraphBarTabs(paragraph) || !ShouldRenderNativeEmptyParagraphLineBox(paragraph)) return;
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, "implicit body paragraph after table", identity: body);
        RenderNativeEmptyParagraph(flow, paragraph, CreateNativeParagraphStyle(paragraph, defaults, fonts), defaults, fonts);
    }
}
