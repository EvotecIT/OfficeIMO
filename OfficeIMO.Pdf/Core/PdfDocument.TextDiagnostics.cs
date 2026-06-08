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
        if (_loadedPdf is not null) {
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
                AnalyzeListItems(list.RichItems, options, defaultFont, diagnostics, location);
                break;
            case NumberedListBlock list:
                AnalyzeListItems(list.RichItems, options, defaultFont, diagnostics, location);
                break;
            case TableBlock table:
                AnalyzeTable(table, options, defaultFont, diagnostics, "PdfTableCell", location);
                break;
            case PanelParagraphBlock panel:
                AddRuns(diagnostics, panel.Runs, options, defaultFont, "PdfPanel", location);
                break;
            case TextFieldBlock textField:
                AddFormWidgetText(diagnostics, textField.Value, "PdfTextField", location, fieldName: textField.Name);
                break;
            case ChoiceFieldBlock choiceField:
                AnalyzeChoiceField(choiceField, diagnostics, location);
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

    private static void AnalyzeListItems(IReadOnlyList<PdfListItem> items, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
            PdfListItem item = items[itemIndex];
            string itemLocation = AppendLocation(locationPrefix, "PdfListItem[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]");
            AddText(diagnostics, item.Marker, options, defaultFont, "PdfListMarker", AppendLocation(itemLocation, "Marker"));
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
                    AnalyzeTableFormField(field, diagnostics, cellLocation, rowIndex, cellIndex);
                }
            }
        }
    }

    private static void AnalyzeTableCaption(TableBlock table, PdfOptions options, PdfStandardFont defaultFont, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        string? caption = table.Style?.Caption;
        if (string.IsNullOrWhiteSpace(caption)) {
            return;
        }

        AddText(diagnostics, caption, options, defaultFont, "PdfTableCaption", AppendLocation(locationPrefix, "PdfTableCaption"));
    }

    private static void AnalyzeChoiceField(ChoiceFieldBlock choiceField, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix) {
        for (int optionIndex = 0; optionIndex < choiceField.Options.Count; optionIndex++) {
            AddFormWidgetText(diagnostics, choiceField.Options[optionIndex], "PdfChoiceFieldOption", AppendLocation(locationPrefix, "Option[" + optionIndex.ToString(CultureInfo.InvariantCulture) + "]"), fieldName: choiceField.Name);
        }

        for (int valueIndex = 0; valueIndex < choiceField.Values.Count; valueIndex++) {
            AddFormWidgetText(diagnostics, choiceField.Values[valueIndex], "PdfChoiceFieldValue", AppendLocation(locationPrefix, "Value[" + valueIndex.ToString(CultureInfo.InvariantCulture) + "]"), fieldName: choiceField.Name);
        }
    }

    private static void AnalyzeTableFormField(PdfTableCellFormField field, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix, int rowIndex, int cellIndex) {
        if (field.Kind == PdfTableCellFormFieldKind.Text) {
            AddFormWidgetText(diagnostics, field.Value, "PdfTableTextField", AppendLocation(locationPrefix, "PdfTableTextField"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
            return;
        }

        for (int optionIndex = 0; optionIndex < field.Options.Count; optionIndex++) {
            AddFormWidgetText(diagnostics, field.Options[optionIndex], "PdfTableChoiceFieldOption", AppendLocation(locationPrefix, "PdfTableChoiceFieldOption[" + optionIndex.ToString(CultureInfo.InvariantCulture) + "]"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
        }

        for (int valueIndex = 0; valueIndex < field.Values.Count; valueIndex++) {
            AddFormWidgetText(diagnostics, field.Values[valueIndex], "PdfTableChoiceFieldValue", AppendLocation(locationPrefix, "PdfTableChoiceFieldValue[" + valueIndex.ToString(CultureInfo.InvariantCulture) + "]"), tableRowIndex: rowIndex, tableColumnIndex: cellIndex, fieldName: field.Name);
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
                    font: font)
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
                case PdfCanvasDrawingItem drawing:
                    AnalyzeDrawing(drawing.Block, options, defaultFont, diagnostics, AppendLocation(locationPrefix, "PdfCanvasDrawing[" + itemIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                    break;
            }
        }
    }

    private static void AnalyzeOptionsText(PdfOptions options, List<PdfTextEncodingDiagnostic> diagnostics, string locationPrefix = "") {
        var seenPageText = new HashSet<string>(StringComparer.Ordinal);
        PdfTextWatermark? watermark = options.TextWatermarkSnapshot;
        if (watermark is not null) {
            PdfStandardFont watermarkFont = PdfStandardFontMapper.GetStyledFont(watermark.Font, watermark.Bold, watermark.Italic);
            AddText(diagnostics, watermark.Text, options, watermarkFont, "PdfTextWatermark", AppendLocation(locationPrefix, "PdfTextWatermark"));
        }

        AnalyzePageTextForVariant(options, pageNumber: 1, diagnostics, seenPageText, locationPrefix);
        AnalyzePageTextForVariant(options, pageNumber: 2, diagnostics, seenPageText, locationPrefix);
        AnalyzePageTextForVariant(options, pageNumber: 3, diagnostics, seenPageText, locationPrefix);
    }

    private static void AnalyzePageTextForVariant(PdfOptions options, int pageNumber, List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, string locationPrefix) {
        if (options.HasHeaderTextContentForPage(pageNumber)) {
            string headerLocation = AppendLocation(locationPrefix, "PdfHeader[page=" + pageNumber.ToString(CultureInfo.InvariantCulture) + "]");
            AddPageText(diagnostics, seenPageText, options.GetHeaderFormatForPage(pageNumber), options, options.HeaderFont, "PdfHeader", headerLocation, pageNumber);
            AddSegments(diagnostics, seenPageText, options.GetHeaderSegmentsForPage(pageNumber), options, options.HeaderFont, "PdfHeader", headerLocation, pageNumber);
            AddZones(diagnostics, seenPageText, options.GetHeaderZonesForPage(pageNumber), options, options.HeaderFont, "PdfHeader", headerLocation, pageNumber);
        }

        if (options.HasFooterTextContentForPage(pageNumber)) {
            string footerLocation = AppendLocation(locationPrefix, "PdfFooter[page=" + pageNumber.ToString(CultureInfo.InvariantCulture) + "]");
            AddPageText(diagnostics, seenPageText, options.GetFooterFormatForPage(pageNumber), options, options.FooterFont, "PdfFooter", footerLocation, pageNumber);
            AddSegments(diagnostics, seenPageText, options.GetFooterSegmentsForPage(pageNumber), options, options.FooterFont, "PdfFooter", footerLocation, pageNumber);
            AddZones(diagnostics, seenPageText, options.GetFooterZonesForPage(pageNumber), options, options.FooterFont, "PdfFooter", footerLocation, pageNumber);
        }
    }

    private static void AddSegments(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, IReadOnlyList<FooterSegment>? segments, PdfOptions options, PdfStandardFont font, string source, string locationPrefix, int pageNumber) {
        if (segments is null) {
            return;
        }

        for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
            FooterSegment? segment = segments[segmentIndex];
            if (segment is null) {
                throw new ArgumentException(source + " segments cannot contain null entries.");
            }

            if (segment.Kind == FooterSegmentKind.Text) {
                AddPageText(diagnostics, seenPageText, segment.Text, options, font, source, AppendLocation(locationPrefix, "Segment[" + segmentIndex.ToString(CultureInfo.InvariantCulture) + "]"), pageNumber);
            }
        }
    }

    private static void AddZones(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, (string? Left, string? Center, string? Right) zones, PdfOptions options, PdfStandardFont font, string source, string locationPrefix, int pageNumber) {
        AddPageText(diagnostics, seenPageText, zones.Left, options, font, source, AppendLocation(locationPrefix, "Left"), pageNumber);
        AddPageText(diagnostics, seenPageText, zones.Center, options, font, source, AppendLocation(locationPrefix, "Center"), pageNumber);
        AddPageText(diagnostics, seenPageText, zones.Right, options, font, source, AppendLocation(locationPrefix, "Right"), pageNumber);
    }

    private static void AddText(List<PdfTextEncodingDiagnostic> diagnostics, string? text, PdfOptions options, PdfStandardFont font, string source, string location, int? pageNumber = null, int? tableRowIndex = null, int? tableColumnIndex = null, string? fieldName = null) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> textDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(text!, options, font, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in textDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber, tableRowIndex, tableColumnIndex, fieldName));
        }
    }

    private static void AddFormWidgetText(List<PdfTextEncodingDiagnostic> diagnostics, string? text, string source, string location, int? tableRowIndex = null, int? tableColumnIndex = null, string? fieldName = null) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        IReadOnlyList<PdfTextEncodingDiagnostic> textDiagnostics = PdfTextDiagnostics.AnalyzeWinAnsiText(text!, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in textDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber: null, tableRowIndex, tableColumnIndex, fieldName));
        }
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

    private static void AddPageText(List<PdfTextEncodingDiagnostic> diagnostics, HashSet<string> seenPageText, string? text, PdfOptions options, PdfStandardFont font, string source, string location, int pageNumber) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        string key = source + "|" + font + "|" + text;
        if (seenPageText.Add(key)) {
            AddText(diagnostics, text, options, font, source, location, pageNumber);
        }
    }

    private static void AddRuns(List<PdfTextEncodingDiagnostic> diagnostics, IEnumerable<TextRun> runs, PdfOptions options, PdfStandardFont defaultFont, string source, string location, int? tableRowIndex = null, int? tableColumnIndex = null) {
        IReadOnlyList<PdfTextEncodingDiagnostic> runDiagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(runs, options, defaultFont, source, location);
        foreach (PdfTextEncodingDiagnostic diagnostic in runDiagnostics) {
            diagnostics.Add(AnnotateDiagnostic(diagnostic, pageNumber: null, tableRowIndex: tableRowIndex, tableColumnIndex: tableColumnIndex));
        }
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
