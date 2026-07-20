using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Analyzes generated document text against the configured PDF text encoding before rendering.
    /// </summary>
    /// <remarks>
    /// Opened byte-backed PDFs do not have generated OfficeIMO blocks to inspect and return an empty result.
    /// </remarks>
    public IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeTextEncoding() {
        if (_source is not null) {
            return Array.Empty<PdfTextEncodingDiagnostic>();
        }

        var diagnostics = new List<PdfTextEncodingDiagnostic>();
        AnalyzeOptionsText(_options, diagnostics);
        AnalyzeBlocks(_blocks, _options, diagnostics, string.Empty);
        return diagnostics.AsReadOnly();
    }

    private bool TryCreateTextEncodingPreflightException(out PdfTextEncodingPreflightException? exception) {
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = AnalyzeTextEncoding();
        if (diagnostics.Count == 0) {
            exception = null;
            return false;
        }

        _options.AddTextDiagnostics(diagnostics);
        exception = new PdfTextEncodingPreflightException(diagnostics);
        return true;
    }

    private static void AnalyzeBlocks(IEnumerable<IPdfBlock> blocks, PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        int blockIndex = 0;
        foreach (IPdfBlock block in blocks) {
            AnalyzeBlock(block, options, diagnostics, AppendLocation(locationPrefix, GetBlockLocation(block, blockIndex)));
            blockIndex++;
        }
    }

    private static void AnalyzeBlock(IPdfBlock block, PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string location) {
        PdfStandardFont defaultFont = PdfStandardFontMapper.GetFontFamily(options.DefaultFont);
        switch (block) {
            case RichParagraphBlock paragraph:
                AddRuns(diagnostics, paragraph.Runs, options, defaultFont, "PdfParagraph", location);
                break;
            case ParagraphBlock paragraph:
                AddText(diagnostics, paragraph.Text, options, defaultFont, "PdfParagraph", location);
                break;
            case HeadingBlock heading:
                AddText(diagnostics, heading.Text, options, GetHeadingFont(heading, options), "PdfHeading", location);
                break;
            case BulletListBlock list:
                AnalyzeListItems(list.RichItems, list.Style ?? options.DefaultListStyleSnapshot, options, defaultFont, diagnostics, location, numbered: false, startNumber: 1);
                break;
            case NumberedListBlock list:
                AnalyzeListItems(list.RichItems, list.Style ?? options.DefaultListStyleSnapshot, options, defaultFont, diagnostics, location, numbered: true, list.StartNumber);
                break;
            case TableBlock table:
                AnalyzeTable(table, options, defaultFont, diagnostics, "PdfTableCell", location);
                break;
            case DeferredTableBlock table:
                AnalyzeDeferredTable(table, options, defaultFont, diagnostics, location);
                break;
            case PanelParagraphBlock panel:
                AddRuns(diagnostics, panel.Runs, options, defaultFont, "PdfPanel", location);
                break;
            case TextFieldBlock textField:
                AddFormWidgetText(diagnostics, textField.Value, options, "PdfTextField", location, fieldName: textField.Name);
                break;
            case ChoiceFieldBlock choiceField:
                AnalyzeChoiceField(choiceField, options, diagnostics, location);
                break;
            case FreeTextAnnotationBlock freeText:
                AddFreeTextAppearanceText(diagnostics, freeText.Contents, options, "PdfFreeTextAnnotation", location);
                break;
            case DrawingBlock drawing:
                AnalyzeDrawing(drawing, options, defaultFont, diagnostics, location);
                break;
            case PdfCanvasBlock canvas:
                AnalyzeCanvasItems(canvas.Items, options, defaultFont, diagnostics, location);
                break;
            case RowBlock row:
                for (int columnIndex = 0; columnIndex < row.Columns.Count; columnIndex++) {
                    AnalyzeBlocks(row.Columns[columnIndex].Blocks, options, diagnostics, AppendLocation(location, "Column[" + columnIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                }
                break;
            case PageBlock page:
                AnalyzeOptionsText(page.Options, diagnostics, location);
                AnalyzeBlocks(page.Blocks, page.Options, diagnostics, location);
                break;
        }
    }

    private static void AnalyzeListItems(IReadOnlyList<PdfListItem> items, PdfListStyle? style, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix, bool numbered, int startNumber) {
        PdfStandardFont markerFont = PdfStandardFontMapper.GetFontFamily(style?.MarkerFont ?? defaultFont);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
            PdfListItem item = items[itemIndex];
            string itemLocation = AppendLocation(locationPrefix, "PdfListItem[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]");
            string marker = item.Marker ??
                (numbered
                    ? (startNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."
                    : "•");
            var markerRuns = new[] {
                new TextRun(
                    marker,
                    bold: style?.MarkerBold == true,
                    italic: style?.MarkerItalic == true,
                    font: markerFont,
                    fontFamily: style?.MarkerFontFamily)
            };
            AddRuns(diagnostics, markerRuns, options, markerFont, "PdfListMarker", AppendLocation(itemLocation, "Marker"));
            AddRuns(diagnostics, item.Runs, options, defaultFont, "PdfListItem", itemLocation);
        }
    }

    private static void AnalyzeTable(TableBlock table, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string source, string locationPrefix) {
        AnalyzeTableCaption(table, options, defaultFont, diagnostics, locationPrefix);
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            IReadOnlyList<PdfTableCell> row = table.Cells[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                PdfTableCell cell = row[cellIndex];
                string cellLocation = AppendLocation(locationPrefix, source + "[" + rowIndex.ToString(CultureInfo.InvariantCulture) + "," + cellIndex.ToString(CultureInfo.InvariantCulture) + "]");
                AddRuns(diagnostics, cell.Runs, options, defaultFont, source, cellLocation, rowIndex, cellIndex);
                foreach (PdfTableCellFormField field in cell.FormFields) {
                    AnalyzeTableFormField(field, options, diagnostics, cellLocation, rowIndex, cellIndex);
                }
            }
        }
    }

    private static void AnalyzeDeferredTable(DeferredTableBlock table, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        string? caption = table.Style?.Caption;
        if (!string.IsNullOrWhiteSpace(caption)) {
            AddText(diagnostics, caption, options, defaultFont, "PdfTableCaption", AppendLocation(locationPrefix, "PdfTableCaption"));
        }

        int rowIndex = 0;
        foreach (PdfTableCell[] row in table.EnumerateRows()) {
            for (int cellIndex = 0; cellIndex < row.Length; cellIndex++) {
                PdfTableCell cell = row[cellIndex];
                string cellLocation = AppendLocation(locationPrefix, "PdfTableCell[" + rowIndex.ToString(CultureInfo.InvariantCulture) + "," + cellIndex.ToString(CultureInfo.InvariantCulture) + "]");
                AddRuns(diagnostics, cell.Runs, options, defaultFont, "PdfTableCell", cellLocation, rowIndex, cellIndex);
                foreach (PdfTableCellFormField field in cell.FormFields) {
                    AnalyzeTableFormField(field, options, diagnostics, cellLocation, rowIndex, cellIndex);
                }
            }

            rowIndex++;
        }
    }

    private static void AnalyzeTableCaption(TableBlock table, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        string? caption = table.Style?.Caption;
        if (string.IsNullOrWhiteSpace(caption)) {
            return;
        }

        AddText(diagnostics, caption, options, defaultFont, "PdfTableCaption", AppendLocation(locationPrefix, "PdfTableCaption"));
    }

    private static void AnalyzeChoiceField(ChoiceFieldBlock choiceField, PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        for (int optionIndex = 0; optionIndex < choiceField.Options.Count; optionIndex++) {
            AddFormWidgetText(diagnostics, choiceField.Options[optionIndex], options, "PdfChoiceFieldOption", AppendLocation(locationPrefix, "Option[" + optionIndex.ToString(CultureInfo.InvariantCulture) + "]"), fieldName: choiceField.Name);
        }

        for (int valueIndex = 0; valueIndex < choiceField.Values.Count; valueIndex++) {
            AddFormWidgetText(diagnostics, choiceField.Values[valueIndex], options, "PdfChoiceFieldValue", AppendLocation(locationPrefix, "Value[" + valueIndex.ToString(CultureInfo.InvariantCulture) + "]"), fieldName: choiceField.Name);
        }
    }

    private static void AnalyzeTableFormField(PdfTableCellFormField field, PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix, int rowIndex, int cellIndex) {
        if (field.Kind == PdfTableCellFormFieldKind.Text) {
            AddFormWidgetText(diagnostics, field.Value, options, "PdfTableTextField", AppendLocation(locationPrefix, "PdfTableTextField"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
            return;
        }

        for (int optionIndex = 0; optionIndex < field.Options.Count; optionIndex++) {
            AddFormWidgetText(diagnostics, field.Options[optionIndex], options, "PdfTableChoiceFieldOption", AppendLocation(locationPrefix, "PdfTableChoiceFieldOption[" + optionIndex.ToString(CultureInfo.InvariantCulture) + "]"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
        }

        for (int valueIndex = 0; valueIndex < field.Values.Count; valueIndex++) {
            AddFormWidgetText(diagnostics, field.Values[valueIndex], options, "PdfTableChoiceFieldValue", AppendLocation(locationPrefix, "PdfTableChoiceFieldValue[" + valueIndex.ToString(CultureInfo.InvariantCulture) + "]"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
        }
    }

    private static void AnalyzeDrawing(DrawingBlock drawing, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        for (int elementIndex = 0; elementIndex < drawing.Drawing.Elements.Count; elementIndex++) {
            OfficeDrawingElement element = drawing.Drawing.Elements[elementIndex];
            if (element is not OfficeDrawingText text || string.IsNullOrEmpty(text.Text)) {
                continue;
            }

            PdfStandardFont font = defaultFont;
            if (!string.IsNullOrWhiteSpace(text.Font.FamilyName) &&
                PdfStandardFontMapper.TryMapFontFamily(text.Font.FamilyName, out PdfStandardFont mappedFont)) {
                font = PdfStandardFontMapper.GetFontFamily(mappedFont);
            }

            var runs = new[] {
                new TextRun(
                    text.Text,
                    bold: text.Font.IsBold,
                    underline: text.Font.IsUnderline,
                    italic: text.Font.IsItalic,
                    font: font,
                    fontFamily: text.Font.FamilyName)
            };
            AddRuns(diagnostics, runs, options, font, "PdfDrawingText", AppendLocation(locationPrefix, "PdfDrawingText[" + elementIndex.ToString(CultureInfo.InvariantCulture) + "]"));
        }
    }

    private static void AnalyzeCanvasItems(IReadOnlyList<PdfCanvasItem> items, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
            PdfCanvasItem item = items[itemIndex];
            switch (item) {
                case PdfCanvasTextItem text:
                    AddRuns(diagnostics, text.Runs, options, defaultFont, "PdfCanvasText", AppendLocation(locationPrefix, "PdfCanvasText[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasTextBoxItem textBox:
                    AddRuns(diagnostics, textBox.Runs, options, defaultFont, "PdfCanvasTextBox", AppendLocation(locationPrefix, "PdfCanvasTextBox[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasFreeTextAnnotationItem freeText:
                    AddFreeTextAppearanceText(diagnostics, freeText.Contents, options, "PdfCanvasFreeTextAnnotation", AppendLocation(locationPrefix, "PdfCanvasFreeTextAnnotation[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasTableItem table:
                    AnalyzeTable(table.Block, options, defaultFont, diagnostics, "PdfCanvasTableCell", AppendLocation(locationPrefix, "PdfCanvasTable[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasClipItem clip:
                    AnalyzeCanvasItems(clip.Items, options, defaultFont, diagnostics, AppendLocation(locationPrefix, "PdfCanvasClip[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasEffectItem effect:
                    AnalyzeCanvasItems(effect.Items, options, defaultFont, diagnostics, AppendLocation(locationPrefix, "PdfCanvasEffect[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
                case PdfCanvasDrawingItem drawing:
                    AnalyzeDrawing(drawing.Block, options, defaultFont, diagnostics, AppendLocation(locationPrefix, "PdfCanvasDrawing[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
            }
        }
    }

    private static void AnalyzeOptionsText(PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix = "") {
        var seenPageText = new HashSet<string>(StringComparer.Ordinal);
        AnalyzeTextWatermark(options.TextWatermarkSnapshot, options, diagnostics, locationPrefix, "PdfTextWatermark");
        AnalyzeTextWatermark(options.FirstPageTextWatermarkSnapshot, options, diagnostics, locationPrefix, "PdfFirstPageTextWatermark");
        AnalyzeTextWatermark(options.EvenPageTextWatermarkSnapshot, options, diagnostics, locationPrefix, "PdfEvenPageTextWatermark");

        AnalyzePageTextForVariant(options, pageNumber: 1, diagnostics, seenPageText, locationPrefix);
        AnalyzePageTextForVariant(options, pageNumber: 2, diagnostics, seenPageText, locationPrefix);
        AnalyzePageTextForVariant(options, pageNumber: 3, diagnostics, seenPageText, locationPrefix);
    }

    private static void AnalyzeTextWatermark(PdfTextWatermark? watermark, PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix, string source) {
        if (watermark is null) {
            return;
        }

        PdfStandardFont watermarkFont = PdfStandardFontMapper.GetStyledFont(watermark.Font, watermark.Bold, watermark.Italic);
        AddText(diagnostics, watermark.Text, options, watermarkFont, source, AppendLocation(locationPrefix, source));
    }

    private static void AnalyzePageTextForVariant(PdfOptions options, int pageNumber, List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, string locationPrefix) {
        if (options.HasHeaderTextContentForPage(pageNumber)) {
            string headerLocation = AppendLocation(locationPrefix, "PdfHeader[page=" + pageNumber.ToString(CultureInfo.InvariantCulture) + "]");
            AddPageText(diagnostics, seenPageText, options.GetHeaderFormatForPage(pageNumber), options, options.HeaderFont, options.HeaderFontFamily, "PdfHeader", headerLocation, pageNumber);
            AddSegments(diagnostics, seenPageText, options.GetHeaderSegmentsForPage(pageNumber), options, options.HeaderFont, options.HeaderFontFamily, "PdfHeader", headerLocation, pageNumber);
            AddZones(diagnostics, seenPageText, options.GetHeaderZonesForPage(pageNumber), options, options.HeaderFont, options.HeaderFontFamily, "PdfHeader", headerLocation, pageNumber);
        }

        if (options.HasFooterTextContentForPage(pageNumber)) {
            string footerLocation = AppendLocation(locationPrefix, "PdfFooter[page=" + pageNumber.ToString(CultureInfo.InvariantCulture) + "]");
            AddPageText(diagnostics, seenPageText, options.GetFooterFormatForPage(pageNumber), options, options.FooterFont, options.FooterFontFamily, "PdfFooter", footerLocation, pageNumber);
            AddSegments(diagnostics, seenPageText, options.GetFooterSegmentsForPage(pageNumber), options, options.FooterFont, options.FooterFontFamily, "PdfFooter", footerLocation, pageNumber);
            AddZones(diagnostics, seenPageText, options.GetFooterZonesForPage(pageNumber), options, options.FooterFont, options.FooterFontFamily, "PdfFooter", footerLocation, pageNumber);
        }
    }

    private static void AddSegments(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, IReadOnlyList<FooterSegment>? segments, PdfOptions options, PdfStandardFont font, string? fontFamily, string source, string locationPrefix, int pageNumber) {
        if (segments is null) {
            return;
        }

        for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
            FooterSegment? segment = segments[segmentIndex];
            if (segment is null) {
                throw new ArgumentException(source + " segments cannot contain null entries.");
            }

            if (segment.Kind == FooterSegmentKind.Text) {
                AddPageText(diagnostics, seenPageText, segment.Text, options, font, fontFamily, source, AppendLocation(locationPrefix, "Segment[" + segmentIndex.ToString(CultureInfo.InvariantCulture) + "]"), pageNumber);
            }
        }
    }

    private static void AddZones(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, (string? Left, string? Center, string? Right) zones, PdfOptions options, PdfStandardFont font, string? fontFamily, string source, string locationPrefix, int pageNumber) {
        AddPageText(diagnostics, seenPageText, zones.Left, options, font, fontFamily, source, AppendLocation(locationPrefix, "Left"), pageNumber);
        AddPageText(diagnostics, seenPageText, zones.Center, options, font, fontFamily, source, AppendLocation(locationPrefix, "Center"), pageNumber);
        AddPageText(diagnostics, seenPageText, zones.Right, options, font, fontFamily, source, AppendLocation(locationPrefix, "Right"), pageNumber);
    }

    private static void AddText(List<PdfTextEncodingDiagnostic> diagnostics, string? text, PdfOptions options, PdfStandardFont font, string source, string location, int? pageNumber = null, int? tableRowIndex = null, int? tableColumnIndex = null, string? fieldName = null) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        if (options.HasDiagnosticsReport) {
            options.AddTextShapingDiagnostics(
                PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text!, source),
                text!,
                ShouldDeferProviderCoveredTextShapingDiagnostics(options, font, text!));
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> textDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(text!, options, font, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in textDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber, tableRowIndex, tableColumnIndex, fieldName));
        }
    }

    private static void AddFormWidgetText(List<PdfTextEncodingDiagnostic> diagnostics, string? text, PdfOptions options, string source, string location, int? tableRowIndex = null, int? tableColumnIndex = null, string? fieldName = null) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> textDiagnostics = AnalyzeFormWidgetText(text!, options, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in textDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber: null, tableRowIndex, tableColumnIndex, fieldName));
        }
    }

    private static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeFormWidgetText(string text, PdfOptions options, string source, string location) {
        IReadOnlyList<PdfTextEncodingDiagnostic> generatedDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(text, options, PdfStandardFont.Helvetica, source, location);
        if (generatedDiagnostics.Count == 0) {
            return Array.Empty<PdfTextEncodingDiagnostic>();
        }

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacksSnapshot;
        if (fallbackSet != null) {
            PdfTextFallbackPlan plan = fallbackSet.PlanText(text, source);
            if (plan.IsFullyCovered) {
                return Array.Empty<PdfTextEncodingDiagnostic>();
            }

            return plan.Diagnostics;
        }

        return generatedDiagnostics;
    }

    private static void AddFreeTextAppearanceText(List<PdfTextEncodingDiagnostic> diagnostics, string? text, PdfOptions options, string source, string location) {
        if (string.IsNullOrEmpty(text) || !options.TryGetEmbeddedStandardFont(PdfStandardFont.Helvetica, out _)) {
            return;
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> textDiagnostics = PdfTextDiagnostics.AnalyzeWinAnsiText(text!, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in textDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber: null, tableRowIndex: null, tableColumnIndex: null));
        }
    }

    private static void AddPageText(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, string? text, PdfOptions options, PdfStandardFont font, string? fontFamily, string source, string location, int pageNumber) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        string key = source + "|" + font + "|" + fontFamily + "|" + text;
        if (seenPageText.Add(key)) {
            PdfStandardFont fallbackFont = PdfStandardFontMapper.GetFontFamily(font);
            var runs = new[] {
                new TextRun(
                    text!,
                    bold: IsBoldStandardFont(font),
                    italic: IsItalicStandardFont(font),
                    font: fallbackFont,
                    fontFamily: fontFamily)
            };
            if (options.HasDiagnosticsReport) {
                AddRunTextShapingDiagnostics(runs, options, fallbackFont, source);
            }

            IReadOnlyList<PdfTextEncodingDiagnostic> runDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(runs, options, fallbackFont, source, location);
            foreach (PdfTextEncodingDiagnostic diagnostic in runDiagnostics) {
                var pageTextDiagnostic = new PdfTextEncodingDiagnostic(
                    diagnostic.Source,
                    diagnostic.Index,
                    diagnostic.CodePoint,
                    diagnostic.Text,
                    diagnostic.IsControlCharacter,
                    diagnostic.Encoding,
                    diagnostic.Remediation,
                    location,
                    runIndex: null,
                    pageNumber: pageNumber);
                diagnostics.Add(pageTextDiagnostic);
            }
        }
    }

    private static void AddRuns(List<PdfTextEncodingDiagnostic> diagnostics, IEnumerable<TextRun> runs, PdfOptions options, PdfStandardFont defaultFont, string source, string location, int? tableRowIndex = null, int? tableColumnIndex = null, int? pageNumber = null) {
        if (options.HasDiagnosticsReport) {
            AddRunTextShapingDiagnostics(runs, options, defaultFont, source);
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> runDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(runs, options, defaultFont, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in runDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber, tableRowIndex, tableColumnIndex));
        }
    }

    private static bool IsBoldStandardFont(PdfStandardFont font) =>
        font is PdfStandardFont.HelveticaBold or
            PdfStandardFont.HelveticaBoldOblique or
            PdfStandardFont.TimesBold or
            PdfStandardFont.TimesBoldItalic or
            PdfStandardFont.CourierBold or
            PdfStandardFont.CourierBoldOblique;

    private static bool IsItalicStandardFont(PdfStandardFont font) =>
        font is PdfStandardFont.HelveticaOblique or
            PdfStandardFont.HelveticaBoldOblique or
            PdfStandardFont.TimesItalic or
            PdfStandardFont.TimesBoldItalic or
            PdfStandardFont.CourierOblique or
            PdfStandardFont.CourierBoldOblique;

    private static void AddRunTextShapingDiagnostics(IEnumerable<TextRun> runs, PdfOptions options, PdfStandardFont defaultFont, string source) {
        foreach (TextRun run in runs) {
            if (run == null || string.Equals(run.Text, "\n", StringComparison.Ordinal) || string.Equals(run.Text, "\t", StringComparison.Ordinal)) {
                continue;
            }

            PdfStandardFont runFont = ResolveDiagnosticRunFont(defaultFont, run);
            options.AddTextShapingDiagnostics(
                PdfTextDiagnostics.AnalyzeAdvancedTextLayout(run.Text, source),
                run.Text,
                ShouldDeferProviderCoveredTextShapingDiagnostics(options, runFont, run.Text));
        }
    }

    private static bool ShouldDeferProviderCoveredTextShapingDiagnostics(PdfOptions options, PdfStandardFont font, string text) {
        if (options.TextShapingProviderSnapshot == null) {
            return false;
        }

        if (HasEmbeddedFontProgram(options, font)) {
            return true;
        }

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacksSnapshot;
        if (fallbackSet == null || string.IsNullOrEmpty(text)) {
            return false;
        }

        PdfTextFallbackPlan plan = fallbackSet.PlanText(text, shapingMode: options.TextShapingModeSnapshot);
        if (!plan.IsFullyCovered) {
            return false;
        }

        foreach (PdfTextFallbackSegment segment in plan.Segments) {
            if (segment.FontIndex >= 0 &&
                segment.FontIndex < fallbackSet.FontSlots.Count &&
                HasEmbeddedFontProgram(options, fallbackSet.FontSlots[segment.FontIndex])) {
                return true;
            }
        }

        return false;
    }

    private static bool HasEmbeddedFontProgram(PdfOptions options, PdfStandardFont font) =>
        (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) && fontProgram != null) ||
        (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) && cffFontProgram != null);

    private static PdfStandardFont ResolveDiagnosticRunFont(PdfStandardFont defaultFont, TextRun run) {
        PdfStandardFont font = run.Font ?? defaultFont;
        if (run.Bold && run.Italic) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: true, italic: true);
        }

        if (run.Bold) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: true, italic: false);
        }

        if (run.Italic) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: false, italic: true);
        }

        return font;
    }

    private static PdfTextEncodingDiagnostic AnnotateDiagnostic(PdfTextEncodingDiagnostic diagnostic, int? pageNumber, int? tableRowIndex, int? tableColumnIndex, string? fieldName = null) {
        PdfTextEncodingDiagnostic annotated = diagnostic;
        if (pageNumber.HasValue) {
            annotated = annotated.WithPageNumber(pageNumber.Value);
        }

        if (tableRowIndex.HasValue && tableColumnIndex.HasValue) {
            annotated = annotated.WithTableCell(tableRowIndex.Value, tableColumnIndex.Value);
        }

        if (!string.IsNullOrWhiteSpace(fieldName)) {
            annotated = annotated.WithFieldName(fieldName!);
        }

        return annotated;
    }

    private static PdfStandardFont GetHeadingFont(HeadingBlock heading, PdfOptions options) {
        PdfHeadingStyle? style = heading.Style ?? options.DefaultHeadingStylesSnapshot?.GetSnapshot(heading.Level);
        return PdfStandardFontMapper.GetStyledFont(options.DefaultFont, style?.Bold ?? true, italic: false);
    }

    private static string GetBlockLocation(IPdfBlock block, int blockIndex) {
        string index = blockIndex.ToString(CultureInfo.InvariantCulture);
        return block switch {
            RichParagraphBlock => "PdfParagraph[" + index + "]",
            ParagraphBlock => "PdfParagraph[" + index + "]",
            HeadingBlock => "PdfHeading[" + index + "]",
            BulletListBlock => "PdfBulletList[" + index + "]",
            NumberedListBlock => "PdfNumberedList[" + index + "]",
            TableBlock => "PdfTable[" + index + "]",
            PanelParagraphBlock => "PdfPanel[" + index + "]",
            TextFieldBlock => "PdfTextField[" + index + "]",
            ChoiceFieldBlock => "PdfChoiceField[" + index + "]",
            FreeTextAnnotationBlock => "PdfFreeTextAnnotation[" + index + "]",
            DrawingBlock => "PdfDrawing[" + index + "]",
            PdfCanvasBlock => "PdfCanvas[" + index + "]",
            RowBlock => "PdfRow[" + index + "]",
            PageBlock => "PdfPage[" + index + "]",
            _ => "PdfBlock[" + index + "]"
        };
    }

    private static string AppendLocation(string prefix, string segment) {
        if (string.IsNullOrWhiteSpace(prefix)) {
            return segment ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(segment)) {
            return prefix;
        }

        return prefix + "." + segment;
    }
}
