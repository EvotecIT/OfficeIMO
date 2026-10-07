using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void RenderNativeElement(INativePdfFlow pdf, WordElement element, WordSection activeSection, Func<WordParagraph, (int Level, string Marker)?> getMarker, IReadOnlyList<int> footnoteNumbers, NativeNoteNumbering footnoteNumbersById, WordToPdfOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap = null, WordElement? nextElement = null) {
            nativeFontMap ??= new NativeFontMap();
            switch (element) {
                case WordParagraph paragraph:
                    RenderNativeParagraph(pdf, paragraph, getMarker(paragraph), getMarker, footnoteNumbers, footnoteNumbersById, options, headingDestinations, nativeDefaults, nativeFontMap, nextElement as WordParagraph);
                    break;
                case WordTableOfContent tableOfContent:
                    RenderNativeTableOfContents(pdf, tableOfContent, tableOfContentsEntries, contentWidth);
                    break;
                case WordTable table:
                    RenderNativeTable(pdf, table, getMarker, footnoteNumbersById, options, contentWidth, nativeDefaults, nativeFontMap);
                    break;
                case WordImage image:
                    RenderNativeImage(pdf, image, options: options, source: "body image");
                    break;
                case WordHyperLink link:
                    RenderNativeHyperLink(pdf, link);
                    break;
                case WordBreak wordBreak:
                    RenderNativeBreak(pdf, wordBreak);
                    break;
                case WordShape shape:
                    RenderNativeShape(pdf, shape);
                    break;
                case WordCoverPage coverPage:
                    RenderNativeCoverPage(pdf, coverPage, activeSection, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, nativeDefaults, nativeFontMap);
                    break;
                case WordStructuredDocumentTag structuredDocumentTag:
                    RenderNativeStructuredDocumentTag(pdf, structuredDocumentTag, activeSection, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, nativeDefaults, nativeFontMap);
                    break;
                case WordWatermark:
                    break;
                case WordEmbeddedDocument:
                    if (options != null) {
                        AddNativeExportWarning(
                            options,
                            "NativeBodyEmbeddedDocumentUnsupported",
                            "body",
                            "Embedded documents in Word body content are not mapped by the OfficeIMO PDF engine yet.");
                    }

                    break;
                default:
                    if (options != null) {
                        AddNativeExportWarning(
                            options,
                            "NativeBodyElementUnsupported",
                            "body",
                            "Word body element '" + element.GetType().Name + "' is not mapped by the OfficeIMO PDF engine yet.");
                    }

                    break;
            }
        }

        private static void RenderNativeBreak(INativePdfFlow pdf, WordBreak wordBreak) {
            if (wordBreak.BreakType == WordBreakType.Page) {
                pdf.PageBreak(preserveEmptyPage: true);
            } else if (wordBreak.BreakType == WordBreakType.Column) {
                pdf.ColumnBreak();
            }
        }

        private static void RenderNativeParagraph(INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker, Func<WordParagraph, (int Level, string Marker)?> getMarker, IReadOnlyList<int> footnoteNumbers, NativeNoteNumbering footnoteNumbersById, WordToPdfOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap, WordParagraph? nextParagraph) {
            if (paragraph == null) {
                return;
            }

            // A formatted section terminator is editable metadata, not a separate blank body line.
            if (WordParagraph.IsSectionMarkOnly(paragraph._paragraph)) {
                if (!string.IsNullOrEmpty(paragraph.Bookmark?.Name)) pdf.Bookmark(paragraph.Bookmark!.Name!);
                return;
            }

            if (TryRenderNativeParagraphFlowBreaks(pdf, paragraph, marker, getMarker, footnoteNumbersById,
                options, headingDestinations, nativeDefaults, nativeFontMap, nextParagraph)) return;

            if (HasNativePageBreakBefore(paragraph)) {
                pdf.PageBreak();
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            WordParagraph? currentChartRun = runs.FirstOrDefault(run =>
                ReferenceEquals(run._run, paragraph._run) && run.Chart != null);
            RecordNativeBodyParagraphDiagnostics(paragraph, options, "body paragraph", mapsCheckBoxes: true, mapsFormFields: true, mapsPictureControls: true, mapsRepeatingSections: true);
            IReadOnlyList<W.SdtRun> checkboxControls = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFieldControls = GetNativeFormFieldControls(paragraph);
            IReadOnlyList<W.SdtRun> repeatingSectionControls = GetNativeRepeatingSectionControls(paragraph);

            if (!string.IsNullOrEmpty(paragraph.Bookmark?.Name)) {
                pdf.Bookmark(paragraph.Bookmark!.Name!);
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out string? textBoxFallbackText);
            if (textBox != null) {
                if (marker is { Marker.Length: > 0 }) {
                    PdfCore.PdfParagraphStyle markerStyle = CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap);
                    ApplyNativeInlineListIndent(paragraph, markerStyle);
                    ApplyNativeInlineListMarkerAlignment(paragraph, marker.Value.Marker, markerStyle, nativeDefaults, nativeFontMap);
                    markerStyle.SpacingAfter = 0D;
                    pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, marker,
                        Array.Empty<WordParagraph>(), false, string.Empty, Array.Empty<int>(), footnoteNumbersById, options, nativeDefaults, nativeFontMap,
                        inlineMarkerColumnWidth: Math.Max(0D, -markerStyle.FirstLineIndent)),
                        ResolveNativeParagraphAlign(paragraph, allowJustify: false), style: markerStyle);
                }
                RenderNativeTextBox(pdf, textBox, getMarker, footnoteNumbersById, options, nativeDefaults, nativeFontMap, textBoxFallbackText);
                return;
            }

            PdfCore.PdfAlign objectAlign = ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap);
            if (marker is { Marker.Length: 0 }) ApplyNativeMarkerlessListIndent(paragraph, style);
            bool inlineListMarker = marker is { Marker.Length: > 0 };
            if (inlineListMarker) {
                ApplyNativeInlineListIndent(paragraph, style);
                ApplyNativeInlineListMarkerAlignment(paragraph, marker!.Value.Marker, style, nativeDefaults, nativeFontMap);
            }
            bool hasEquationContent = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0;
            string content = hasEquationContent
                ? AppendNativeTextWithEquation(paragraph.Text, paragraph)
                : paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
            bool hasRenderableRuns = runs.Any(run => IsNativeRenderableTextRun(run, paragraph));
            bool shouldRenderDirectContent = ShouldRenderNativeDirectText(paragraph, runs, content);
            string renderContent = hasRenderableRuns || shouldRenderDirectContent ? content : string.Empty;
            List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, footnoteNumbers, footnoteNumbersById);
            bool needsAnchorLine = style.AnchoredCanvas != null && !hasRenderableRuns &&
                string.IsNullOrEmpty(renderContent) && marker == null && paragraphFootnoteNumbers.Count == 0;
            if (ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph)) {
                style.SpacingAfter = 0D;
            }
            WordShape? currentShape = runs.Where(run => ReferenceEquals(run._run, paragraph._run))
                .Select(run => run.Shape).FirstOrDefault(shape => shape != null);
            bool objectOnly = !hasRenderableRuns && string.IsNullOrEmpty(renderContent) &&
                marker == null && paragraphFootnoteNumbers.Count == 0 && checkboxControls.Count == 0 &&
                formFieldControls.Count == 0 && repeatingSectionControls.Count == 0;
            NativeObjectParagraphSpacing? objectSpacing = objectOnly
                ? new NativeObjectParagraphSpacing(pdf, style, MeasureNativeEmptyParagraphLineHeight(paragraph, nativeDefaults, nativeFontMap))
                : null;
            OfficeDrawing? directChartDrawing = PrepareNativeChart(currentChartRun?.Chart, options, "body paragraph chart");
            List<OfficeDrawing> runChartDrawings = PrepareNativeRunCharts(runs, options, currentChartRun);
            bool renderedChart = directChartDrawing != null || runChartDrawings.Count > 0;
            if (directChartDrawing != null) {
                RenderNativeFlowObject(pdf, objectSpacing, flow => flow.Drawing(directChartDrawing, objectAlign,
                    spacingBefore: objectOnly ? 0D : 2D,
                    spacingAfter: 0D));
            }

            int groupedImageCount = 0;
            bool renderedFlowObject = RenderNativeParagraphShapeGroups(pdf, paragraph, runs, objectAlign, options, style, ref groupedImageCount, objectSpacing);
            if (currentShape != null) {
                renderedFlowObject |= RenderNativeShape(pdf, currentShape, spacingAfter: objectOnly ? 0D : 6D, paragraphSpacing: objectSpacing);
            }

            renderedFlowObject |= RenderNativeParagraphImages(pdf, paragraph, runs, objectAlign, options, style, groupedImageCount, objectSpacing);
            needsAnchorLine = style.AnchoredCanvas != null &&
                !runs.Any(run => IsNativeRenderableTextRun(run, paragraph) && !string.IsNullOrWhiteSpace(run.Text)) &&
                string.IsNullOrWhiteSpace(renderContent) && marker == null && paragraphFootnoteNumbers.Count == 0;
            RenderNativeRunCharts(pdf, runChartDrawings, objectAlign,
                objectOnly ? 0D : 2D, 0D, objectSpacing);
            if (objectSpacing?.Complete() == true) needsAnchorLine = false;

            if (!needsAnchorLine && marker == null &&
                paragraphFootnoteNumbers.Count == 0 &&
                IsNativeHorizontalRuleParagraph(paragraph, runs, renderContent) &&
                CreateNativeHorizontalRuleStyle(paragraph, style) is { } horizontalRuleStyle) {
                pdf.HR(style: horizontalRuleStyle);
                return;
            }

            if (!needsAnchorLine && !hasRenderableRuns && string.IsNullOrEmpty(renderContent) && marker == null &&
                paragraphFootnoteNumbers.Count == 0 && checkboxControls.Count == 0 && formFieldControls.Count == 0 && repeatingSectionControls.Count == 0) {
                // A flow object already occupies this paragraph's line. A real empty
                // paragraph following it still contributes its own visible mark.
                if (renderedFlowObject || renderedChart) {
                    return;
                }
                RenderNativeEmptyParagraph(
                    pdf,
                    paragraph,
                    style,
                    nativeDefaults,
                    nativeFontMap);
                return;
            }

            PdfCore.PdfAlign align = ResolveNativeParagraphAlign(paragraph);
            PdfCore.PdfColor? defaultColor = ResolveNativeParagraphDefaultColor(paragraph);
            int headingLevel = GetHeadingLevel(paragraph);
            PdfCore.PdfColor? headingColor = GetNativeHeadingColor(headingLevel, defaultColor);
            PdfCore.PdfHorizontalRuleStyle? topBorderRuleStyle = marker == null ? CreateNativeTopBorderRuleStyle(paragraph, style) : null;
            PdfCore.PdfParagraphStyle paragraphStyle = topBorderRuleStyle == null ? style : style.Clone();
            if (topBorderRuleStyle != null) {
                paragraphStyle.SpacingBefore = 0;
                pdf.HR(style: topBorderRuleStyle);
            }

            if (!needsAnchorLine && headingLevel > 0 && marker == null) {
                if (paragraph._paragraph != null &&
                    string.IsNullOrEmpty(paragraph.Bookmark?.Name) &&
                    headingDestinations.TryGetValue(paragraph._paragraph, out string? generatedDestinationName)) {
                    pdf.Bookmark(generatedDestinationName);
                }

                string headingText = GetNativeHeadingText(renderContent, runs, paragraph, nativeFontMap, hasEquationContent);
                RenderNativeHeading(pdf, headingLevel, headingText,
                    builder => AddNativeParagraphContent(builder, paragraph, null, runs, hasRenderableRuns,
                        renderContent, paragraphFootnoteNumbers, footnoteNumbersById, options, nativeDefaults, nativeFontMap),
                    objectAlign, headingColor, paragraph, paragraphStyle, nativeDefaults, nativeFontMap);
                if (CreateNativeBottomBorderRuleStyle(paragraph, paragraphStyle) is { } headingRuleStyle) {
                    pdf.HR(style: headingRuleStyle);
                }

                RenderNativeFormFields(pdf, formFieldControls, objectAlign);
                RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
                RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
                return;
            }

            PdfCore.PdfPanelStyle? panelStyle = CreateNativeParagraphPanelStyle(paragraph, paragraphStyle);
            if (panelStyle != null) {
                paragraphStyle = JoinNativeAdjacentParagraphShading(nextParagraph, paragraphStyle, panelStyle, nativeDefaults);
                pdf.PanelParagraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, footnoteNumbersById, options, nativeDefaults, nativeFontMap, needsAnchorLine,
                        inlineMarkerColumnWidth: inlineListMarker ? Math.Max(0D, -paragraphStyle.FirstLineIndent) : null);
                }, panelStyle, align, defaultColor, paragraphStyle);
                RenderNativeFormFields(pdf, formFieldControls, objectAlign);
                RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
                RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
                return;
            }

            PdfCore.PdfHorizontalRuleStyle? bottomBorderRuleStyle = marker == null ? CreateNativeBottomBorderRuleStyle(paragraph, paragraphStyle) : null;
            if (bottomBorderRuleStyle != null && ReferenceEquals(paragraphStyle, style)) {
                paragraphStyle = style.Clone();
            }

            if (bottomBorderRuleStyle != null) {
                paragraphStyle.SpacingAfter = 0;
            }

            if (needsAnchorLine || hasRenderableRuns || !string.IsNullOrEmpty(renderContent) || marker != null || paragraphFootnoteNumbers.Count > 0) {
                pdf.Paragraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, footnoteNumbersById, options, nativeDefaults, nativeFontMap, needsAnchorLine,
                        inlineMarkerColumnWidth: inlineListMarker ? Math.Max(0D, -paragraphStyle.FirstLineIndent) : null);
                }, align, defaultColor, paragraphStyle);
            }

            if (bottomBorderRuleStyle != null) {
                pdf.HR(style: bottomBorderRuleStyle);
            }

            RenderNativeFormFields(pdf, formFieldControls, objectAlign);
            RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
            RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
        }

        private static void RenderNativeEmptyParagraph(
            INativePdfFlow pdf,
            WordParagraph paragraph,
            PdfCore.PdfParagraphStyle style,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            if (!ShouldRenderNativeEmptyParagraphLineBox(paragraph)) {
                // A hidden mark contributes neither a line nor its paragraph spacing.
                return;
            }

            double height = MeasureNativeEmptyParagraphLineHeight(
                paragraph,
                nativeDefaults,
                nativeFontMap);
            if (height > 0D) {
                pdf.ParagraphSpacingBefore(style.SpacingBefore);
                pdf.Spacer(height);
                pdf.ParagraphSpacingAfter(style.SpacingAfter ?? nativeDefaults.ParagraphSpacingAfter);
            }
        }

        private static bool ShouldRenderNativeEmptyParagraphLineBox(WordParagraph paragraph) {
            // Empty text still has a paragraph mark. Its visibility, rather than
            // the presence of direct spacing or run formatting, determines the line.
            bool hiddenMark = ReadNativeOnOff(paragraph._paragraph?.ParagraphProperties?
                .ParagraphMarkRunProperties?.GetFirstChild<W.Vanish>()) ??
                GetNativeParagraphStyleDefaults(paragraph).Hidden ?? false;
            return !hiddenMark;
        }

        private static double MeasureNativeEmptyParagraphLineHeight(
            WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = ResolveNativeParagraphLayoutFontSize(paragraph, nativeDefaults, styleDefaults);
            double lineHeight = ResolveNativeParagraphLineHeight(
                paragraph,
                fontSize,
                nativeDefaults,
                styleDefaults,
                nativeFontMap);
            double height = fontSize * lineHeight;
            return double.IsNaN(height) || double.IsInfinity(height) ? 0D : Math.Max(0D, height);
        }

        private static double MeasureNativeEmptyParagraphHeight(WordParagraph paragraph, PdfCore.PdfParagraphStyle style,
            NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) =>
            style.SpacingBefore + MeasureNativeEmptyParagraphLineHeight(paragraph, nativeDefaults, nativeFontMap) +
            (style.SpacingAfter ?? nativeDefaults.ParagraphSpacingAfter);

        private static void RenderNativeFormFields(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> formFieldControls, PdfCore.PdfAlign align) {
            for (int index = 0; index < formFieldControls.Count; index++) {
                W.SdtRun formField = formFieldControls[index];
                double spacingBefore = index == 0 ? 0D : 2D;
                if (IsNativeDatePickerControl(formField)) {
                    pdf.TextField(
                        GetNativeContentControlFieldName(formField, index, "WordDatePicker"),
                        width: 150D,
                        height: 20D,
                        value: GetNativeDatePickerValue(formField),
                        align: align,
                        fontSize: 10D,
                        spacingBefore: spacingBefore,
                        spacingAfter: 4D);
                    continue;
                }

                IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                string? value = GetNativeChoiceFieldValue(formField, options);
                if (options.Count == 0 || string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                string fallbackPrefix = formField.SdtProperties?.Elements<W.SdtContentComboBox>().Any() == true
                    ? "WordComboBox"
                    : "WordDropDownList";
                pdf.ChoiceField(
                    GetNativeContentControlFieldName(formField, index, fallbackPrefix),
                    options,
                    value,
                    width: 150D,
                    height: 20D,
                    align: align,
                    fontSize: 10D,
                    spacingBefore: spacingBefore,
                    spacingAfter: 4D,
                    isComboBox: true);
            }
        }

        private static void RenderNativeCheckBoxes(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> checkboxControls, PdfCore.PdfAlign align) {
            for (int index = 0; index < checkboxControls.Count; index++) {
                W.SdtRun checkbox = checkboxControls[index];
                pdf.CheckBox(
                    GetNativeCheckBoxFieldName(checkbox, index),
                    IsNativeCheckBoxChecked(checkbox),
                    size: 12D,
                    align: align,
                    spacingBefore: index == 0 ? 0D : 2D,
                    spacingAfter: 4D);
            }
        }

        private static void RenderNativeRepeatingSections(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> repeatingSectionControls, PdfCore.PdfAlign align, PdfCore.PdfColor? color) {
            foreach (W.SdtRun repeatingSection in repeatingSectionControls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(repeatingSection)) {
                    if (string.IsNullOrWhiteSpace(itemText)) {
                        continue;
                    }

                    pdf.Paragraph(builder => builder.Text(NormalizeNativeDirectText(itemText)), align, color);
                }
            }
        }

        private static IReadOnlyList<string> GetNativeRepeatingSectionItems(W.SdtRun repeatingSection) {
            var items = new List<string>();
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> itemElements = repeatingSection.SdtContentRun?.ChildElements
                .Where(element => element.LocalName == "repeatingSectionItem") ??
                Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>();

            foreach (DocumentFormat.OpenXml.OpenXmlElement item in itemElements) {
                string text = string.Concat(item.Descendants<W.Text>().Select(value => value.Text));
                if (!string.IsNullOrWhiteSpace(text)) {
                    items.Add(text);
                }
            }

            if (items.Count == 0) {
                string text = GetNativeSdtText(repeatingSection) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text)) {
                    items.Add(text);
                }
            }

            return items;
        }

        private readonly record struct NativeTableCellEmbeddedContent(
            IReadOnlyList<PdfCore.PdfTableCellCheckBox> CheckBoxes,
            IReadOnlyList<PdfCore.PdfTableCellFormField> FormFields,
            IReadOnlyList<PdfCore.PdfTableCellImage> Images,
            IReadOnlyDictionary<DocumentFormat.OpenXml.OpenXmlElement, PdfCore.PdfTextRun> InlineImages);

        private static NativeTableCellEmbeddedContent CreateNativeTableCellEmbeddedContent(WordTableCell cell, WordToPdfOptions? options) {
            List<PdfCore.PdfTableCellCheckBox>? checkBoxes = null;
            List<PdfCore.PdfTableCellFormField>? formFields = null;
            List<PdfCore.PdfTableCellImage>? images = null;
            var inlineImages = new Dictionary<DocumentFormat.OpenXml.OpenXmlElement, PdfCore.PdfTextRun>();
            int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
            if (imageLimit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
            foreach (WordParagraph paragraph in EnumerateNativeTableCellParagraphs(cell)) {
                int imageCount = 0;
                W.Paragraph? paragraphElement = paragraph._paragraph;
                bool supportsInlineImages = paragraphElement != null &&
                    WordEquation.GetOccurrences(paragraph._document, paragraphElement).Count == 0;
                void AddImage(WordImage image) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    if (++imageCount > imageLimit)
                        throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                    if (supportsInlineImages && paragraphElement != null && image._Image?.Inline != null &&
                        !image._Image.Ancestors<W.SdtRun>().Any(IsNativePictureControl) &&
                        ReferenceEquals(image._Image.Ancestors<W.TextBoxContent>().FirstOrDefault(),
                            paragraphElement.Ancestors<W.TextBoxContent>().FirstOrDefault())) {
                        if (TryCreateNativeCellInlineImage(image, out PdfCore.PdfTextRun? inline))
                            inlineImages[image._Image] = inline!;
                        return;
                    }
                    images ??= new List<PdfCore.PdfTableCellImage>();
                    AddNativeTableCellImage(images, image);
                }
                foreach (WordImage image in EnumerateNativeParagraphImages(paragraph, options?.CancellationToken ?? default))
                    AddImage(image);

                if (paragraph._paragraph == null) {
                    continue;
                }

                foreach (W.SdtRun control in paragraph._paragraph.Descendants<W.SdtRun>()) {
                    if (IsNativeCheckBoxControl(control)) {
                        checkBoxes ??= new List<PdfCore.PdfTableCellCheckBox>();
                        checkBoxes.Add(new PdfCore.PdfTableCellCheckBox(
                            GetNativeCheckBoxFieldName(control, checkBoxes.Count, "WordTableCheckBox"),
                            IsNativeCheckBoxChecked(control),
                            size: 12D));
                        continue;
                    }

                    if (IsNativeSupportedFormFieldContentControl(control)) {
                        formFields ??= new List<PdfCore.PdfTableCellFormField>();
                        AddNativeTableCellFormField(formFields, control);
                        continue;
                    }

                }
            }

            return new NativeTableCellEmbeddedContent(
                checkBoxes ?? (IReadOnlyList<PdfCore.PdfTableCellCheckBox>)Array.Empty<PdfCore.PdfTableCellCheckBox>(),
                formFields ?? (IReadOnlyList<PdfCore.PdfTableCellFormField>)Array.Empty<PdfCore.PdfTableCellFormField>(),
                images ?? (IReadOnlyList<PdfCore.PdfTableCellImage>)Array.Empty<PdfCore.PdfTableCellImage>(), inlineImages);
        }

        private static IEnumerable<WordImage> EnumerateNativeParagraphImages(WordParagraph paragraph, CancellationToken cancellationToken, int textBoxDepth = 0) {
            if (paragraph._paragraph == null) {
                foreach (WordImage image in paragraph.EnumerateImages()) yield return image;
                yield break;
            }
            foreach (WordParagraph imageRun in GetNativeRuns(paragraph)) {
                cancellationToken.ThrowIfCancellationRequested();
                W.Run? run = imageRun._run;
                if (run == null) continue;
                if (run.Ancestors<W.DeletedRun>().Any() || run.Ancestors<W.MoveFromRun>().Any()) continue;
                if (IsNativeHiddenTextRun(imageRun, paragraph)) continue;
                if (run.Ancestors<W.SdtRun>().Any(IsNativePictureControl)) continue;
                W.TextBoxContent? currentTextBox = paragraph._paragraph.Ancestors<W.TextBoxContent>().FirstOrDefault();
                foreach (WordImage image in imageRun.EnumerateImages()) {
                    DocumentFormat.OpenXml.OpenXmlElement imageElement = image._vmlShape ?? (DocumentFormat.OpenXml.OpenXmlElement)image._Image;
                    if (ReferenceEquals(imageElement.Ancestors<W.TextBoxContent>().FirstOrDefault(), currentTextBox))
                        yield return image;
                }
                if (textBoxDepth >= 8) continue;
                IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> visibleChildren = imageRun._visibleRunSourceChildren ??
                    (IEnumerable<DocumentFormat.OpenXml.OpenXmlElement>)run.ChildElements;
                foreach (W.TextBoxContent content in visibleChildren.SelectMany(child => child.Descendants<W.TextBoxContent>())
                    .Where(content => ReferenceEquals(content.Ancestors<W.TextBoxContent>().FirstOrDefault(), currentTextBox))) {
                    foreach (W.Paragraph inner in content.Descendants<W.Paragraph>()
                        .Where(inner => ReferenceEquals(inner.Ancestors<W.TextBoxContent>().FirstOrDefault(), content))) {
                        var innerParagraph = new WordParagraph(paragraph._document, inner);
                        foreach (WordImage image in EnumerateNativeParagraphImages(innerParagraph, cancellationToken, textBoxDepth + 1))
                            yield return image;
                    }
                }
            }
            foreach (W.SdtRun control in GetNativePictureControls(paragraph)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (control.Ancestors<W.DeletedRun>().Any() || control.Ancestors<W.MoveFromRun>().Any()) continue;
                var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph, control);
                WordImage? image = pictureParagraph.PictureControl?.Image;
                if (image != null && !IsNativeHiddenImageContent(image, paragraph)) yield return image;
            }
        }

        private static void AddNativeTableCellFormField(List<PdfCore.PdfTableCellFormField> formFields, W.SdtRun formField) {
            if (IsNativeDatePickerControl(formField)) {
                formFields.Add(PdfCore.PdfTableCellFormField.TextField(
                    GetNativeContentControlFieldName(formField, formFields.Count, "WordTableDatePicker"),
                    GetNativeDatePickerValue(formField),
                    width: 150D,
                    height: 20D,
                    fontSize: 10D));
                return;
            }

            IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
            string? value = GetNativeChoiceFieldValue(formField, options);
            if (options.Count == 0 || string.IsNullOrWhiteSpace(value)) {
                return;
            }

            string fallbackPrefix = formField.SdtProperties?.Elements<W.SdtContentComboBox>().Any() == true
                ? "WordTableComboBox"
                : "WordTableDropDownList";
            formFields.Add(PdfCore.PdfTableCellFormField.ChoiceField(
                GetNativeContentControlFieldName(formField, formFields.Count, fallbackPrefix),
                options,
                value,
                width: 150D,
                height: 20D,
                fontSize: 10D,
                isComboBox: true));
        }

        private static IEnumerable<WordParagraph> EnumerateNativeTableCellParagraphs(WordTableCell cell, int tableNestingDepth = 0) {
            if (IsNativeHorizontalMergeContinuation(cell) || IsNativeVerticalMergeContinuation(cell)) {
                yield break;
            }

            foreach (WordElement element in EnumerateNativeTableCellElements(cell)) {
                if (element is WordParagraph paragraph) {
                    yield return paragraph;
                    continue;
                }

                if (element is not WordTable nestedTable) {
                    continue;
                }

                int nestedDepth = tableNestingDepth + 1;
                EnsureNativeTableDepth(nestedDepth);
                foreach (WordTableRow nestedRow in nestedTable.Rows) {
                    foreach (WordTableCell nestedCell in nestedRow.Cells) {
                        foreach (WordParagraph nestedParagraph in EnumerateNativeTableCellParagraphs(nestedCell, nestedDepth)) {
                            yield return nestedParagraph;
                        }
                    }
                }
            }
        }

        private static void AddNativeTableCellImage(List<PdfCore.PdfTableCellImage> images, WordImage image) {
            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!TryPrepareNativePdfImageBytes(bytes, out byte[] preparedBytes, out _)) {
                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            images.Add(new PdfCore.PdfTableCellImage(preparedBytes, width, height, CreateNativeImageStyle()));
        }

        private static void AddNativeParagraphContent(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph paragraph,
            (int Level, string Marker)? marker,
            IReadOnlyList<WordParagraph> runs,
            bool hasRenderableRuns,
            string content,
            IReadOnlyList<int> paragraphFootnoteNumbers,
            NativeNoteNumbering footnoteNumbersById,
            WordToPdfOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap,
            bool needsAnchorLine = false,
            double? inlineMarkerColumnWidth = null) {
            if (needsAnchorLine) {
                // An image-only paragraph still participates in pagination and decoration.
                builder.FontSize(ResolveNativeParagraphFontSize(paragraph,
                    nativeDefaults, GetNativeParagraphStyleDefaults(paragraph))).LineBreak();
                return;
            }
            if (marker != null && !string.IsNullOrEmpty(marker.Value.Marker)) {
                builder.Text(new string(' ', Math.Max(0, marker.Value.Level - 1) * 2));
                WordDocumentTraversal.ListInfo? markerInfo = WordDocumentTraversal.GetListInfo(paragraph);
                double leadingMarkerOffset = 0D;
                double trailingMarkerOffset = 0D;
                bool useAlignedMarkerColumn = false;
                if (markerInfo.HasValue) {
                    NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
                    PdfCore.PdfTextRun measuredMarker = CreateNativeListMarkerTextRun(marker.Value.Marker,
                        paragraph, textStyle, nativeFontMap, includeSuffix: false);
                    if (inlineMarkerColumnWidth.HasValue && nativeFontMap.MeasureText(measuredMarker) is { } measuredWidth) {
                        trailingMarkerOffset = markerInfo.Value.LevelSuffix switch {
                            WordListLevelSuffix.Nothing => 0D,
                            WordListLevelSuffix.Space => Math.Max(0D,
                                (nativeFontMap.MeasureText(CreateNativeListMarkerTextRun(marker.Value.Marker,
                                    paragraph, textStyle, nativeFontMap)) ?? measuredWidth) - measuredWidth),
                            _ => Math.Max(0D, inlineMarkerColumnWidth.Value - measuredWidth)
                        };
                        useAlignedMarkerColumn = true;
                    } else if (inlineMarkerColumnWidth.HasValue &&
                        (markerInfo.Value.LevelJustification == WordListLevelAlignment.Right ||
                         markerInfo.Value.LevelJustification == WordListLevelAlignment.Center)) {
                        double markerFontSize = markerInfo.Value.MarkerFontSize ?? textStyle.FontSize ?? nativeDefaults.FontSize;
                        NativeTextSpacing markerSpacing = ResolveNativeListMarkerTextSpacing(markerInfo.Value, textStyle.ListMarkerTextSpacing);
                        double markerWidth = EstimateNativeListMarkerWidth(marker.Value.Marker, markerFontSize, markerSpacing);
                        double markerColumnWidth = Math.Max(markerWidth, Math.Max(0D, inlineMarkerColumnWidth.Value));
                        leadingMarkerOffset = markerInfo.Value.LevelJustification == WordListLevelAlignment.Right
                            ? Math.Max(0D, markerColumnWidth - markerWidth)
                            : Math.Max(0D, (markerColumnWidth - markerWidth) / 2D);
                        double suffixWidth = markerInfo.Value.LevelSuffix == WordListLevelSuffix.Space
                            ? EstimateNativeListMarkerWidth(" ", markerFontSize, markerSpacing)
                            : 0D;
                        trailingMarkerOffset = Math.Max(0D, markerColumnWidth - leadingMarkerOffset - markerWidth) + suffixWidth;
                        useAlignedMarkerColumn = true;
                        AddNativeInlineListMarkerSpacer(builder, leadingMarkerOffset);
                    }
                    ApplyNativeTextStyle(builder, textStyle);
                    NativeTextSpacing resolvedMarkerSpacing = ResolveNativeListMarkerTextSpacing(markerInfo.Value, textStyle.ListMarkerTextSpacing);
                    builder.HorizontalTextScaling(resolvedMarkerSpacing.WidthPercentage ?? 100D);
                    builder.CharacterSpacing(resolvedMarkerSpacing.CharacterSpacing ?? 0D);
                    builder.Bold(markerInfo.Value.MarkerBold ?? textStyle.Bold);
                    builder.Italic(markerInfo.Value.MarkerItalic ?? textStyle.Italic);
                    if (markerInfo.Value.MarkerFontSize.HasValue) builder.FontSize(markerInfo.Value.MarkerFontSize.Value);
                    if (ParseNativeColor(markerInfo.Value.MarkerColorHex) is { } markerColor) builder.Color(markerColor);
                    string? markerFontFamily = ResolveNativeListMarkerFontFamily(markerInfo.Value, marker.Value.Marker, textStyle, nativeFontMap);
                    if (!string.IsNullOrWhiteSpace(markerFontFamily)) builder.FontFamily(markerFontFamily!);
                    else if (ResolveNativeListMarkerFont(markerInfo.Value, marker.Value.Marker, textStyle) is { } markerFont) builder.Font(markerFont);
                }
                builder.Text(marker.Value.Marker);
                if (markerInfo.HasValue) ResetNativeTextStyle(builder);
                if (useAlignedMarkerColumn) {
                    AddNativeInlineListMarkerSpacer(builder, trailingMarkerOffset);
                } else {
                    builder.Text(ResolveNativeInlineListMarkerSuffix(markerInfo?.LevelSuffix));
                }
            }

            IReadOnlyList<WordTabStop> tabStops = GetNativeParagraphEffectiveTabStops(paragraph);
            int tabIndex = 0;
            bool hasEquationContent = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0;
            if (hasRenderableRuns && !hasEquationContent) {
                foreach (WordParagraph run in runs) {
                    if (run.IsImage && run.Image != null) {
                        continue;
                    }

                    if (IsNativeHiddenTextRun(run, paragraph)) {
                        continue;
                    }

                    if (IsNativeTextWrappingBreak(run) && string.IsNullOrEmpty(run.Text)) {
                        AddNativeRun(builder, "\n", run, paragraph, tabStops, ref tabIndex, options, nativeDefaults, nativeFontMap);
                        continue;
                    }

                    AddNativeRun(builder, run, paragraph, tabStops, ref tabIndex, options, nativeDefaults, nativeFontMap);
                }

                string? supplementalText = GetNativeSupplementalTextAfterRuns(content, runs);
                if (!string.IsNullOrEmpty(supplementalText)) {
                    AddNativeText(builder, supplementalText!, paragraph, tabStops, ref tabIndex, nativeDefaults, nativeFontMap);
                }
            } else if (hasEquationContent) {
                AddNativeEquationContent(builder, paragraph, tabStops, ref tabIndex, options, nativeDefaults, nativeFontMap);
            } else if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !IsNativeHiddenTextRun(paragraph) && !string.IsNullOrEmpty(paragraph.Hyperlink.Text)) {
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
                ApplyNativeTextStyle(builder, style);
                AddNativeHyperLinkRun(builder, paragraph.Hyperlink.Text, paragraph.Hyperlink, tabStops, ref tabIndex, style);
                ResetNativeTextStyle(builder);
            } else {
                AddNativeText(builder, content, paragraph, tabStops, ref tabIndex, nativeDefaults, nativeFontMap);
            }

            builder.Runs(CreateNativeNoteReferenceRuns(paragraph, paragraphFootnoteNumbers, footnoteNumbersById,
                nativeDefaults, nativeFontMap));
        }

        private static void AddNativeInlineListMarkerSpacer(PdfCore.PdfParagraphBuilder builder, double width) {
            if (width > 0.01D) {
                builder.InlineBox(width, 0.01D, borderWidth: 0D);
            }
        }

        private static bool RenderNativeParagraphImages(INativePdfFlow pdf, WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, PdfCore.PdfAlign align, WordToPdfOptions? options, PdfCore.PdfParagraphStyle anchorStyle, int groupedImageCount, NativeObjectParagraphSpacing? paragraphSpacing) {
            int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
            if (imageLimit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
            var positions = paragraph._paragraph!.Descendants().Select((element, index) => (element, index))
                .ToDictionary(pair => pair.element, pair => pair.index);
            var images = new List<(WordImage Image, int Position)>();
            void Add(WordImage? image, DocumentFormat.OpenXml.OpenXmlElement? container) {
                if (image == null) return;
                // Group rendering owns its child images and their local coordinates.
                if (image._vmlShape?.Ancestors<DocumentFormat.OpenXml.Vml.Group>().Any() == true) return;
                if (groupedImageCount + images.Count >= imageLimit)
                    throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                // DrawingML images retain their exact position; VML wrappers use their containing run.
                int position = image._Image != null && positions.TryGetValue(image._Image, out int drawingPosition) ? drawingPosition
                    : container != null && positions.TryGetValue(container, out int containerPosition) ? containerPosition : int.MaxValue;
                images.Add((image, position));
            }
            foreach (W.SdtRun control in GetNativePictureControls(paragraph)) {
                var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, control);
                WordImage? image = pictureParagraph.PictureControl?.Image;
                if (image != null && !IsNativeHiddenImageContent(image, paragraph)) Add(image, control);
            }
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (WordImage image in run.EnumerateImages()) Add(image, run._run);
            }
            var anchoredCanvas = new PdfCore.PdfPageCanvas();
            bool renderedFlowObject = false;
            if (anchorStyle.AnchoredCanvas != null) anchoredCanvas.AddItems(anchorStyle.AnchoredCanvas.Items);
            foreach (var image in images.OrderBy(item => item.Position)) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                renderedFlowObject |= RenderNativeImage(pdf, image.Image, align, options, "body paragraph image", anchorStyle, anchoredCanvas, paragraphSpacing);
            }
            if (anchoredCanvas.Items.Count > 0)
                anchorStyle.AnchoredCanvas = new PdfCore.PdfCanvasBlock(anchoredCanvas.Items);
            return renderedFlowObject;
        }

        private static List<OfficeDrawing> PrepareNativeRunCharts(IReadOnlyList<WordParagraph> runs, WordToPdfOptions? options, WordParagraph? currentChartRun) {
            var drawings = new List<OfficeDrawing>();
            foreach (WordParagraph run in runs) {
                // Only the selected view has already been dispatched. Other views
                // can share its source run while owning different drawings.
                if (ReferenceEquals(run, currentChartRun)) {
                    continue;
                }

                OfficeDrawing? drawing = PrepareNativeChart(run.Chart, options, "body paragraph chart run");
                if (drawing != null) {
                    drawings.Add(drawing);
                }
            }
            return drawings;
        }

        private static void RenderNativeRunCharts(INativePdfFlow pdf, IReadOnlyList<OfficeDrawing> drawings, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, NativeObjectParagraphSpacing? paragraphSpacing) {
            for (int index = 0; index < drawings.Count; index++) {
                OfficeDrawing drawing = drawings[index];
                double before = index == 0 ? spacingBefore : 2D;
                double after = index == drawings.Count - 1 ? spacingAfter : 0D;
                RenderNativeFlowObject(pdf, paragraphSpacing,
                    flow => flow.Drawing(drawing, align, spacingBefore: before, spacingAfter: after));
            }
        }

        private static string? GetNativeSupplementalTextAfterRuns(string content, IReadOnlyList<WordParagraph> runs) {
            if (string.IsNullOrEmpty(content)) {
                return null;
            }

            var renderedText = new StringBuilder();
            foreach (WordParagraph run in runs) {
                if (run.IsImage || string.IsNullOrEmpty(run.Text)) {
                    continue;
                }

                renderedText.Append(run.Text);
            }

            if (renderedText.Length == 0) {
                return content;
            }

            string emittedText = renderedText.ToString();
            if (content.Length <= emittedText.Length ||
                !content.StartsWith(emittedText, StringComparison.Ordinal)) {
                return null;
            }

            return content.Substring(emittedText.Length);
        }

        private static bool IsNativeTextWrappingBreak(WordParagraph run) =>
            run.IsBreak && run.Break?.BreakType != WordBreakType.Page;

        private static WordTextBox? GetNativeParagraphTextBox(WordParagraph paragraph, out string? fallbackText) {
            WordTextBox? textBox = FindNativeParagraphTextBox(paragraph);
            fallbackText = GetNativeTextBoxPlainText(paragraph, textBox);
            return textBox;
        }

        private static WordTextBox? FindNativeParagraphTextBox(WordParagraph paragraph) {
            return GetNativeRuns(paragraph).Select(run => run.TextBox).FirstOrDefault(textBox => textBox != null);
        }

        private static string? GetNativeParagraphTextBoxPlainText(WordParagraph paragraph) {
            return GetNativeTextBoxPlainText(paragraph, FindNativeParagraphTextBox(paragraph));
        }

        private static string? GetNativeTextBoxPlainText(WordParagraph paragraph, WordTextBox? textBox) {
            if (textBox?.Content == null) {
                return null;
            }

            string textBoxText = ResolveNativeBuiltInPropertyPlaceholders(
                paragraph._document,
                string.Join("\n", GetNativeTextBoxParagraphs(textBox)
                    .Select(inner => inner.Text.TrimEnd('\r', '\n'))));
            return string.IsNullOrWhiteSpace(textBoxText) ? null : textBoxText;
        }

        private static void RenderNativeTextBox(INativePdfFlow pdf, WordTextBox textBox, Func<WordParagraph, (int Level, string Marker)?> getMarker, NativeNoteNumbering footnoteNumbersById, WordToPdfOptions? options, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap, string? fallbackText = null) {
            IReadOnlyList<WordParagraph> paragraphs = GetNativeTextBoxParagraphs(textBox);
            if (paragraphs.Count == 0 && !string.IsNullOrWhiteSpace(fallbackText)) {
                PdfCore.PdfPanelStyle fallbackStyle = CreateNativeTextBoxPanelStyle(textBox);
                pdf.PanelParagraph(builder => builder.Text(NormalizeNativeDirectText(fallbackText)), fallbackStyle, PdfCore.PdfAlign.Left);
                return;
            }

            if (paragraphs.Count == 0) {
                return;
            }

            PdfCore.PdfPanelStyle style = CreateNativeTextBoxPanelStyle(textBox);
            pdf.Panel(contentBuilder => {
                for (int index = 0; index < paragraphs.Count; index++) {
                    WordParagraph paragraph = paragraphs[index];
                    List<WordParagraph> runs = GetNativeRuns(paragraph);
                    bool hasEquationContent = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0;
                    string content = hasEquationContent
                        ? AppendNativeTextWithEquation(paragraph.Text, paragraph)
                        : paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                    bool hasRenderableRuns = runs.Any(run => IsNativeRenderableTextRun(run, paragraph));
                    string renderContent = hasRenderableRuns || ShouldRenderNativeDirectText(paragraph, runs, content) ? content : string.Empty;
                    List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                    PdfCore.PdfParagraphStyle paragraphStyle = CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap);
                    if (getMarker(paragraph) is { } paragraphMarker) {
                        ApplyNativeInlineListIndent(paragraph, paragraphStyle);
                        ApplyNativeInlineListMarkerAlignment(paragraph, paragraphMarker.Marker, paragraphStyle, nativeDefaults, nativeFontMap);
                    }
                    paragraphStyle.SpacingBefore = 0D;
                    paragraphStyle.SpacingAfter = 0D;
                    contentBuilder.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, getMarker(paragraph),
                        runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, footnoteNumbersById, options, nativeDefaults, nativeFontMap,
                        inlineMarkerColumnWidth: Math.Max(0D, -paragraphStyle.FirstLineIndent)),
                        ResolveNativeParagraphAlign(paragraph, allowJustify: false), style: paragraphStyle);
                }
            }, style);
        }

        private static IReadOnlyList<WordParagraph> GetNativeTextBoxParagraphs(WordTextBox textBox) {
            W.TextBoxContent? content = textBox.Content;
            IReadOnlyList<WordParagraph> directParagraphs = content == null
                ? Array.Empty<WordParagraph>()
                : content.Descendants<W.Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<W.TextBoxContent>().FirstOrDefault(), content))
                    .Select(paragraph => {
                        W.Run? firstRun = paragraph.Descendants<W.Run>()
                            .FirstOrDefault(run => ReferenceEquals(run.Ancestors<W.Paragraph>().FirstOrDefault(), paragraph));
                        return firstRun == null
                            ? new WordParagraph(textBox.Document, paragraph)
                            : new WordParagraph(textBox.Document, paragraph, firstRun);
                    })
                    .ToList();
            if (HasNativeRenderableTextBoxText(directParagraphs) || directParagraphs.Any(paragraph => paragraph.IsListItem)) {
                return directParagraphs;
            }

            IReadOnlyList<WordParagraph> elementParagraphs = CollapseNativeParagraphElements(textBox.Elements)
                .OfType<WordParagraph>()
                .ToList();
            return elementParagraphs.GroupBy(paragraph => paragraph._paragraph).Select(group => group.First()).ToList();
        }

        private static bool HasNativeRenderableTextBoxText(IEnumerable<WordParagraph> paragraphs) {
            foreach (WordParagraph paragraph in paragraphs) {
                List<WordParagraph> runs = GetNativeRuns(paragraph);
                if (runs.Count == 0 && !IsNativeHiddenTextRun(paragraph) &&
                    WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph).IsVisible &&
                    !paragraph._paragraph.Descendants<W.FieldChar>().Any() &&
                    !string.IsNullOrWhiteSpace(paragraph.Text)) {
                    return true;
                }

                if (runs.Any(run => IsNativeRenderableTextRun(run, paragraph) && !string.IsNullOrWhiteSpace(run.Text))) {
                    return true;
                }
            }

            return false;
        }

        private static PdfCore.PdfPanelStyle CreateNativeTextBoxPanelStyle(WordTextBox textBox) {
            var style = new PdfCore.PdfPanelStyle {
                BorderColor = PdfCore.PdfColor.Black,
                BorderWidth = 0.75D,
                PaddingX = 6D,
                PaddingY = 4D,
                SpacingAfter = 6D,
                Align = MapNativeTextBoxBoxAlign(textBox.HorizontalAlignment)
            };

            double maxWidth = ConvertNativeEmusToPoints(textBox.Width);
            if (maxWidth > 0D) {
                style.MaxWidth = maxWidth;
            }

            return style;
        }

        private static PdfCore.PdfAlign MapNativeTextBoxBoxAlign(WordTextBoxHorizontalAlignment alignment) {
            switch (alignment) {
                case WordTextBoxHorizontalAlignment.Center:
                    return PdfCore.PdfAlign.Center;
                case WordTextBoxHorizontalAlignment.Right:
                case WordTextBoxHorizontalAlignment.Outside:
                    return PdfCore.PdfAlign.Right;
                default:
                    return PdfCore.PdfAlign.Left;
            }
        }

        private static void RenderNativeHeading(INativePdfFlow pdf, int level, string text, Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align, PdfCore.PdfColor? color, WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            PdfCore.PdfHeadingStyle style = CreateNativeWordHeadingStyle(level, paragraph, paragraphStyle, nativeDefaults, nativeFontMap);
            string normalizedText = NormalizeNativeDirectText(text);
            if (string.IsNullOrWhiteSpace(normalizedText)) {
                return;
            }

            pdf.Heading(level, normalizedText, build, align, color, style);
        }

        private static string GetNativeHeadingText(string content, IReadOnlyList<WordParagraph> runs, WordParagraph paragraph, NativeFontMap nativeFontMap, bool hasEquationContent) {
            // Equation content is assembled from the Open XML child order, including field
            // results. Ordinary text runs alone omit the math node in a mixed heading.
            if (hasEquationContent) {
                return ApplyNativeTextTransform(NormalizeNativeDirectText(content), paragraph, nativeFontMap: nativeFontMap);
            }

            var builder = new StringBuilder();
            foreach (WordParagraph run in runs) {
                if (run.IsImage) {
                    continue;
                }

                if (IsNativeHiddenTextRun(run, paragraph)) {
                    continue;
                }

                if (IsNativeTextWrappingBreak(run)) {
                    builder.Append(' ');
                    if (string.IsNullOrEmpty(run.Text)) {
                        continue;
                    }
                }

                if (!string.IsNullOrEmpty(run.Text)) {
                    builder.Append(ApplyNativeTextTransform(run.Text, run, paragraph, nativeFontMap: nativeFontMap));
                }
            }

            string runText = NormalizeNativeDirectText(builder.ToString());
            if (!string.IsNullOrWhiteSpace(runText)) {
                return runText;
            }

            return ApplyNativeTextTransform(NormalizeNativeDirectText(content), paragraph, nativeFontMap: nativeFontMap);
        }

        private static bool IsNativeRenderableTextRun(WordParagraph run, WordParagraph? fallback = null) =>
            !run.IsImage &&
            !string.IsNullOrEmpty(run.Text) &&
            !IsNativeHiddenTextRun(run, fallback);

        private static bool ShouldRenderNativeDirectText(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, string content) =>
            runs.Count == 0 &&
            !string.IsNullOrEmpty(content) &&
            WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph).IsVisible &&
            !paragraph._paragraph.Descendants<W.FieldChar>().Any() &&
            !IsNativeHiddenTextRun(paragraph);

        private static string NormalizeNativeDirectText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return string.Empty;
            }

            return text!
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
        }

        private static PdfCore.PdfHeadingStyle CreateNativeWordHeadingStyle(int level, WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = level switch {
                1 => 16D,
                2 => 13D,
                _ => 12D
            };

            var style = new PdfCore.PdfHeadingStyle {
                AnchoredCanvas = paragraphStyle.AnchoredCanvas,
                FontSize = fontSize,
                LineHeight = 1.18D,
                SpacingBefore = level == 1 ? 24D : 10D,
                SpacingAfter = level == 1 ? 5D : 4D,
                Bold = false,
                ApplySpacingBeforeAtTop = true,
                KeepWithNext = true
            };
            if (ResolveNativeHeadingDeclaredFontSize(paragraph, styleDefaults) is { } declaredFontSize) {
                style.FontSize = declaredFontSize;
            }

            if (ResolveNativeParagraphLineSpacing(paragraph, styleDefaults, nativeDefaults).Value.HasValue && paragraphStyle.LineHeight.HasValue) {
                style.FontSize = ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults);
                style.LineHeight = paragraphStyle.LineHeight.Value;
                style.LineSpacing = paragraphStyle.LineSpacing;
            }

            if (HasNativeHeadingDeclaredSpacingBefore(paragraph, styleDefaults)) {
                style.SpacingBefore = paragraphStyle.SpacingBefore;
            }

            if (HasNativeHeadingDeclaredSpacingAfter(paragraph, styleDefaults)) {
                style.SpacingAfter = paragraphStyle.SpacingAfter;
            }

            style.KeepWithNext = ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? styleDefaults.KeepWithNext ?? true;
            string? headingFontFamily = ResolveNativeParagraphStyleFontFamily(paragraph._document, paragraph.StyleId);
            if (nativeFontMap.TryGetNamedFontFamily(headingFontFamily, out string? registeredHeadingFamily)) {
                style.FontFamily = registeredHeadingFamily;
            }
            if (nativeFontMap.TryGetFontSlot(headingFontFamily, out PdfCore.PdfStandardFont headingFont)) {
                style.Font = headingFont;
            }

            return style;
        }

        private static double? ResolveNativeHeadingDeclaredFontSize(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) {
            if (paragraph.FontSizePoints.HasValue && paragraph.FontSizePoints.Value > 0D) {
                return paragraph.FontSizePoints.Value;
            }

            if (styleDefaults.FontSize.HasValue && styleDefaults.FontSize.Value > 0D) {
                return styleDefaults.FontSize.Value;
            }

            return null;
        }

        private static bool HasNativeHeadingDeclaredSpacingBefore(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.LineSpacingBeforePoints.HasValue || styleDefaults.SpacingBefore.HasValue;

        private static bool HasNativeHeadingDeclaredSpacingAfter(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.LineSpacingAfterPoints.HasValue || styleDefaults.SpacingAfter.HasValue;

        private static void AddNativeRun(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph run,
            WordParagraph paragraphStyleFallback,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            WordToPdfOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            AddNativeRun(builder, run.Text, run, paragraphStyleFallback, tabStops, ref tabIndex, options, nativeDefaults, nativeFontMap);
        }

        private static void AddNativeRun(
            PdfCore.PdfParagraphBuilder builder,
            string text,
            WordParagraph run,
            WordParagraph paragraphStyleFallback,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            WordToPdfOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            if (string.IsNullOrEmpty(text) || IsNativeHiddenTextRun(run, paragraphStyleFallback)) {
                return;
            }

            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraphStyleFallback, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            style = style with { FontSize = style.FontSize ?? nativeFontMap.DefaultFontSize ?? nativeDefaults.FontSize };
            ApplyNativeTextStyle(builder, style);

            if (run.IsHyperLink && run.Hyperlink != null) {
                AddNativeHyperLinkRun(builder, ApplyNativeTextTransform(text, run, paragraphStyleFallback, nativeFontMap: nativeFontMap), run.Hyperlink, tabStops, ref tabIndex, style);
            } else {
                AddNativeRunText(builder, ApplyNativeTextTransform(text, run, paragraphStyleFallback, nativeFontMap: nativeFontMap), tabStops, ref tabIndex);
            }

            ResetNativeTextStyle(builder);
        }

        private static void AddNativeEquationContent(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            WordToPdfOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            foreach (WordEquationContentSegment segment in GetNativeVisibleEquationContentSegments(paragraph)) {
                string visibleText = GetNativeEquationSegmentText(segment);
                if (string.IsNullOrEmpty(visibleText)) continue;
                WordParagraph sourceRun = segment.CreateSourceParagraph(paragraph._document, paragraph._paragraph, paragraph);
                AddNativeRun(builder, visibleText, sourceRun, paragraph, tabStops, ref tabIndex, options, nativeDefaults, nativeFontMap);
            }
        }

        private static string GetNativeEquationSegmentText(WordEquationContentSegment segment) {
            if (segment.Equation != null) return segment.Equation.Text;
            if (segment.Text != null) return segment.Text;
            return segment.IsRunArtifact &&
                (segment.ArtifactElement is W.Break || segment.ArtifactElement is W.CarriageReturn)
                    ? "\n"
                    : string.Empty;
        }

        private static bool IsNativeHiddenTextRun(WordParagraph paragraph, WordParagraph? fallback = null) {
            WordParagraph styleSource = fallback ?? paragraph;
            W.RunProperties? runProperties = GetNativeRunProperties(paragraph);
            NativeCharacterStyleDefaults characterStyleDefaults = GetNativeCharacterStyleDefaults(paragraph._document, runProperties);
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(styleSource);
            return ReadNativeOnOff(runProperties?.GetFirstChild<W.Vanish>()) ??
                   characterStyleDefaults.Hidden ??
                   styleDefaults.Hidden ??
                   false;
        }

        private static OfficeTextDecorationStyle? MapNativeUnderlineStyle(W.Underline? underline) {
            if (underline == null) return null;
            W.UnderlineValues value = underline.Val?.Value ?? W.UnderlineValues.Single;
            if (value == W.UnderlineValues.None) return OfficeTextDecorationStyle.None;
            if (value == W.UnderlineValues.Words) return OfficeTextDecorationStyle.Words;
            if (value == W.UnderlineValues.Double || value == W.UnderlineValues.WavyDouble) return OfficeTextDecorationStyle.Double;
            if (value == W.UnderlineValues.Dotted || value == W.UnderlineValues.DottedHeavy) return OfficeTextDecorationStyle.Dotted;
            if (value == W.UnderlineValues.Wave || value == W.UnderlineValues.WavyHeavy) return OfficeTextDecorationStyle.Wavy;
            if (value == W.UnderlineValues.Dash || value == W.UnderlineValues.DashedHeavy ||
                value == W.UnderlineValues.DashLong || value == W.UnderlineValues.DashLongHeavy ||
                value == W.UnderlineValues.DotDash || value == W.UnderlineValues.DashDotHeavy ||
                value == W.UnderlineValues.DotDotDash || value == W.UnderlineValues.DashDotDotHeavy) {
                return OfficeTextDecorationStyle.Dashed;
            }

            return OfficeTextDecorationStyle.Single;
        }

        private static OfficeTextDecorationStyle? MapNativeStrikeStyle(DocumentFormat.OpenXml.OpenXmlElement? runProperties) {
            bool? strike = ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>());
            bool? doubleStrike = ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>());
            if (!strike.HasValue && !doubleStrike.HasValue) return null;
            if (doubleStrike == true) return OfficeTextDecorationStyle.Double;
            if (strike == true) return OfficeTextDecorationStyle.Single;
            return OfficeTextDecorationStyle.None;
        }

        private static PdfCore.PdfTextBaseline MapNativeTextBaseline(W.VerticalPositionValues? baseline) =>
            baseline == W.VerticalPositionValues.Superscript
                ? PdfCore.PdfTextBaseline.Superscript
                : baseline == W.VerticalPositionValues.Subscript
                    ? PdfCore.PdfTextBaseline.Subscript
                    : PdfCore.PdfTextBaseline.Normal;

        private static W.RunProperties? GetNativeRunProperties(WordParagraph paragraph) =>
            paragraph.IsHyperLink ? paragraph.Hyperlink?._runProperties : paragraph._runProperties;

        private static PdfCore.PdfColor? ResolveNativeParagraphDefaultColor(WordParagraph paragraph) {
            string? paragraphMarkColor = paragraph._paragraph?
                .ParagraphProperties?
                .ParagraphMarkRunProperties?
                .GetFirstChild<W.Color>()?
                .Val?
                .Value;
            return ParseNativeColor(paragraphMarkColor) ??
                   ParseNativeColor(GetNativeParagraphStyleDefaults(paragraph).ColorHex);
        }

        private static PdfCore.PdfStandardFont? ResolveNativeTextRunFont(WordParagraph paragraph, WordParagraph? fallback, NativeCharacterStyleDefaults characterStyleDefaults, NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults, NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            if (TryResolveNativeDirectRunFont(paragraph, nativeFontMap, out PdfCore.PdfStandardFont font) ||
                (fallback != null && TryResolveNativeDirectRunFont(fallback, nativeFontMap, out font)) ||
                TryResolveNativeMappedFont(characterStyleDefaults.FontFamily, nativeFontMap, out font) ||
                TryResolveNativeMappedFont(styleDefaults.FontFamily, nativeFontMap, out font) ||
                TryResolveNativeMappedFont(tableRunStyleDefaults.FontFamily, nativeFontMap, out font) ||
                (nativeFontMap?.UsePdfDefaultForDocumentDefaultFont != true &&
                 TryResolveNativeMappedFont(nativeDefaults.FontFamily, nativeFontMap, out font))) {
                return font;
            }

            return null;
        }

        private static bool TryResolveNativeDirectRunFont(WordParagraph paragraph, NativeFontMap? nativeFontMap, out PdfCore.PdfStandardFont font) =>
            TryResolveNativeMappedFont(paragraph.FontFamily, nativeFontMap, out font) ||
            TryResolveNativeMappedFont(paragraph.FontFamilyHighAnsi, nativeFontMap, out font) ||
            TryResolveNativeMappedFont(paragraph.FontFamilyEastAsia, nativeFontMap, out font) ||
            TryResolveNativeMappedFont(paragraph.FontFamilyComplexScript, nativeFontMap, out font);

        private static bool TryResolveNativeMappedFont(string? familyName, NativeFontMap? nativeFontMap, out PdfCore.PdfStandardFont font) =>
            (nativeFontMap != null && nativeFontMap.TryGetFontSlot(familyName, out font)) ||
            PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out font);

        private static string? ResolveNativeTextRunFontFamily(
            WordParagraph paragraph,
            WordParagraph? fallback,
            NativeCharacterStyleDefaults characterStyleDefaults,
            NativeParagraphStyleDefaults styleDefaults,
            NativeTableRunStyleDefaults tableRunStyleDefaults,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap? nativeFontMap) {
            if (nativeFontMap == null) {
                return null;
            }

            foreach (string? familyName in new[] {
                paragraph.FontFamily,
                paragraph.FontFamilyHighAnsi,
                paragraph.FontFamilyEastAsia,
                paragraph.FontFamilyComplexScript,
                fallback?.FontFamily,
                fallback?.FontFamilyHighAnsi,
                fallback?.FontFamilyEastAsia,
                fallback?.FontFamilyComplexScript,
                characterStyleDefaults.FontFamily,
                styleDefaults.FontFamily,
                tableRunStyleDefaults.FontFamily,
                nativeFontMap.UsePdfDefaultForDocumentDefaultFont ? null : nativeDefaults.FontFamily
            }) {
                if (string.IsNullOrWhiteSpace(familyName)) {
                    continue;
                }

                if (nativeFontMap.TryGetNamedFontFamily(familyName, out string? registeredFamilyName)) {
                    return registeredFamilyName;
                }

                // A standard-family match at a higher-precedence source must stop
                // lower-precedence named defaults from overriding the run.
                if (TryResolveNativeMappedFont(familyName, nativeFontMap, out _)) {
                    return null;
                }
            }

            return null;
        }

        private static bool TryGetNativeRunColor(W.RunProperties? runProperties, out PdfCore.PdfColor? color) {
            W.Color? value = runProperties?.GetFirstChild<W.Color>();
            if (value == null) {
                color = null;
                return false;
            }

            color = ParseNativeColor(value.Val?.Value);
            return true;
        }

        private static bool TryGetNativeRunHighlight(W.RunProperties? runProperties, out PdfCore.PdfColor? color) {
            W.Highlight? value = runProperties?.GetFirstChild<W.Highlight>();
            if (value == null) {
                color = null;
                return false;
            }

            color = MapNativeHighlight(value.Val?.Value);
            return true;
        }

        private static void AddNativeRunText(PdfCore.PdfParagraphBuilder builder, string text, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => builder.Text(value),
                () => builder.LineBreak(),
                () => {
                    AddNativeTab(builder, tabStops, currentTabIndex);
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static void AddNativeTextSegments(string text, Action<string> addText, Action addLineBreak, Action addTab, Action resetTabs) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            var buffer = new StringBuilder();
            for (int index = 0; index < text.Length; index++) {
                char ch = text[index];
                if (ch == '\r') {
                    if (index + 1 < text.Length && text[index + 1] == '\n') {
                        continue;
                    }

                    Flush();
                    addLineBreak();
                    resetTabs();
                    continue;
                }

                if (ch == '\n') {
                    Flush();
                    addLineBreak();
                    resetTabs();
                    continue;
                }

                if (ch == '\t') {
                    Flush();
                    addTab();
                    continue;
                }

                buffer.Append(ch);
            }

            Flush();

            void Flush() {
                if (buffer.Length == 0) {
                    return;
                }

                addText(buffer.ToString());
                buffer.Length = 0;
            }
        }

        private static void AddNativeTab(PdfCore.PdfParagraphBuilder builder, IReadOnlyList<WordTabStop> tabStops, int tabIndex) {
            if (tabIndex < tabStops.Count) {
                WordTabStop tabStop = tabStops[tabIndex];
                builder.Tab(MapNativeTabLeader(tabStop.Leader), MapNativeTabAlignment(tabStop.Alignment));
                return;
            }

            builder.Tab();
        }

        private static void AddNativeHyperLinkRun(PdfCore.PdfParagraphBuilder builder, string text, WordHyperLink hyperlink, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex, NativeResolvedTextStyle? style = null) {
            Uri? uri = hyperlink.Uri;
            string? linkUri = uri != null && uri.IsAbsoluteUri ? uri.AbsoluteUri : null;
            string? bookmarkName = linkUri != null || string.IsNullOrWhiteSpace(hyperlink.Anchor) ? null : hyperlink.Anchor;
            if (linkUri == null && bookmarkName == null) {
                AddNativeRunText(builder, text, tabStops, ref tabIndex);
                return;
            }

            string? contents = GetNativeHyperLinkContents(hyperlink);
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => {
                    if (style.HasValue) {
                        builder.Runs(new[] { CreateNativeHyperLinkTextRun(value, linkUri, bookmarkName, contents, style.Value) });
                    } else if (linkUri != null) {
                        builder.Link(value, linkUri, contents: contents);
                    } else {
                        builder.LinkToBookmark(value, bookmarkName!, contents: contents);
                    }
                },
                () => builder.LineBreak(),
                () => {
                    AddNativeTab(builder, tabStops, currentTabIndex);
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static PdfCore.PdfTextRun CreateNativeHyperLinkTextRun(
            string text,
            string? linkUri,
            string? bookmarkName,
            string? contents,
            NativeResolvedTextStyle style) =>
            style.TextSpacing.ApplyTo(new PdfCore.PdfTextRun(
                text,
                bold: style.Bold,
                underline: style.Underline || linkUri != null || bookmarkName != null,
                color: style.Color,
                italic: style.Italic,
                strike: style.Strike,
                fontSize: style.FontSize,
                font: style.Font,
                linkUri: linkUri,
                linkContents: contents,
                baseline: style.Baseline,
                linkDestinationName: bookmarkName,
                backgroundColor: style.BackgroundColor,
                fontFamily: style.FontFamily,
                underlineStyle: style.Underline || linkUri != null || bookmarkName != null
                    ? style.UnderlineStyle == OfficeTextDecorationStyle.None
                        ? OfficeTextDecorationStyle.Single
                        : style.UnderlineStyle
                    : OfficeTextDecorationStyle.None,
                strikeStyle: style.StrikeStyle));

        private static string? GetNativeHyperLinkContents(WordHyperLink hyperlink) =>
            string.IsNullOrWhiteSpace(hyperlink.Tooltip) ? null : hyperlink.Tooltip;

        private static void AddNativeHyperLinkRun(PdfCore.PdfParagraphBuilder builder, WordHyperLink hyperlink) {
            int tabIndex = 0;
            AddNativeHyperLinkRun(builder, hyperlink.Text, hyperlink, Array.Empty<WordTabStop>(), ref tabIndex);
        }

        private static void RenderNativeHyperLink(INativePdfFlow pdf, WordHyperLink link) {
            if (link == null || string.IsNullOrEmpty(link.Text)) {
                return;
            }

            pdf.Paragraph(builder => AddNativeHyperLinkRun(builder, link));
        }

    }
}
