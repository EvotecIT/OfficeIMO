using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void RenderNativeElement(INativePdfFlow pdf, WordElement element, WordSection activeSection, Func<WordParagraph, (int Level, string Marker)?> getMarker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap = null, bool renderSpacingOnlyEmptyParagraphLineBox = false, WordElement? nextElement = null) {
            nativeFontMap ??= new NativeFontMap();
            switch (element) {
                case WordParagraph paragraph:
                    RenderNativeParagraph(pdf, paragraph, getMarker(paragraph), footnoteNumbers, footnoteNumbersById, options, headingDestinations, nativeDefaults, nativeFontMap, renderSpacingOnlyEmptyParagraphLineBox, nextElement as WordParagraph);
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
            if (wordBreak.BreakType == W.BreakValues.Page) {
                pdf.PageBreak();
            }
        }

        private static void RenderNativeParagraph(INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap, bool renderSpacingOnlyEmptyParagraphLineBox, WordParagraph? nextParagraph) {
            if (paragraph == null) {
                return;
            }

            if (HasNativePageBreakBefore(paragraph)) {
                pdf.PageBreak();
            }

            if (paragraph.IsPageBreak) {
                pdf.PageBreak();
                return;
            }

            RecordNativeBodyParagraphDiagnostics(paragraph, options, "body paragraph", mapsCheckBoxes: true, mapsFormFields: true, mapsPictureControls: true, mapsRepeatingSections: true);
            IReadOnlyList<W.SdtRun> checkboxControls = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFieldControls = GetNativeFormFieldControls(paragraph);
            IReadOnlyList<W.SdtRun> repeatingSectionControls = GetNativeRepeatingSectionControls(paragraph);

            if (!string.IsNullOrEmpty(paragraph.Bookmark?.Name)) {
                pdf.Bookmark(paragraph.Bookmark!.Name!);
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out string? textBoxFallbackText);
            if (textBox != null) {
                RenderNativeTextBox(pdf, textBox, footnoteNumbersById, options, nativeDefaults, nativeFontMap, textBoxFallbackText);
                return;
            }

            PdfCore.PdfAlign objectAlign = ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            RenderNativeChart(pdf, paragraph.Chart, objectAlign, options, "body paragraph chart");

            if (paragraph.Shape != null) {
                RenderNativeShape(pdf, paragraph.Shape);
            }

            if (paragraph.Image != null) {
                RenderNativeImage(pdf, paragraph.Image, objectAlign, options, "body paragraph image");
            }

            WordImage? pictureControlImage = paragraph.PictureControl?.Image;
            if (pictureControlImage != null) {
                RenderNativeImage(pdf, pictureControlImage, objectAlign, options, "body picture control image");
            }

            foreach (W.SdtRun pictureControl in GetNativePictureControls(paragraph)) {
                if (ReferenceEquals(pictureControl, paragraph._stdRun)) {
                    continue;
                }

                var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, pictureControl);
                WordImage? inlinePictureControlImage = pictureParagraph.PictureControl?.Image;
                if (inlinePictureControlImage != null) {
                    RenderNativeImage(pdf, inlinePictureControlImage, objectAlign, options, "body picture control image");
                }
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (paragraph.Image == null) {
                RenderNativeRunImages(pdf, runs, objectAlign, options);
            }

            RenderNativeRunCharts(pdf, runs, objectAlign, options, paragraph._run);

            bool hasEquationContent = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0;
            string content = hasEquationContent
                ? AppendNativeTextWithEquation(paragraph.Text, paragraph)
                : paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
            bool hasRenderableRuns = runs.Any(run => IsNativeRenderableTextRun(run, paragraph));
            bool shouldRenderDirectContent = ShouldRenderNativeDirectText(paragraph, runs, content);
            string renderContent = hasRenderableRuns || shouldRenderDirectContent ? content : string.Empty;
            List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, footnoteNumbers, footnoteNumbersById);
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, nativeDefaults);
            if (ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph)) {
                style.SpacingAfter = 0D;
            }

            if (marker == null &&
                paragraphFootnoteNumbers.Count == 0 &&
                IsNativeHorizontalRuleParagraph(paragraph, runs, renderContent) &&
                CreateNativeHorizontalRuleStyle(paragraph, style) is { } horizontalRuleStyle) {
                pdf.HR(style: horizontalRuleStyle);
                return;
            }

            if (!hasRenderableRuns && string.IsNullOrEmpty(renderContent) && marker == null && paragraphFootnoteNumbers.Count == 0 && checkboxControls.Count == 0 && formFieldControls.Count == 0 && repeatingSectionControls.Count == 0) {
                RenderNativeEmptyParagraph(pdf, paragraph, style, nativeDefaults, renderSpacingOnlyEmptyParagraphLineBox);
                return;
            }

            PdfCore.PdfAlign align = ResolveNativeParagraphAlign(paragraph);
            PdfCore.PdfColor? defaultColor = ParseNativeColor(paragraph.ColorHex);
            int headingLevel = GetHeadingLevel(paragraph);
            PdfCore.PdfColor? headingColor = GetNativeHeadingColor(headingLevel, defaultColor);
            (string? LinkUri, string? LinkDestinationName, string? LinkContents) headingLink = GetNativeHeadingLink(paragraph);
            bool hasHeadingLinkTarget = headingLink.LinkUri != null || headingLink.LinkDestinationName != null;
            PdfCore.PdfHorizontalRuleStyle? topBorderRuleStyle = marker == null ? CreateNativeTopBorderRuleStyle(paragraph, style) : null;
            PdfCore.PdfParagraphStyle paragraphStyle = topBorderRuleStyle == null ? style : style.Clone();
            if (topBorderRuleStyle != null) {
                paragraphStyle.SpacingBefore = 0;
                pdf.HR(style: topBorderRuleStyle);
            }

            if (headingLevel > 0 && marker == null) {
                if (paragraph._paragraph != null &&
                    string.IsNullOrEmpty(paragraph.Bookmark?.Name) &&
                    headingDestinations.TryGetValue(paragraph._paragraph, out string? generatedDestinationName)) {
                    pdf.Bookmark(generatedDestinationName);
                }

                string headingText = GetNativeHeadingText(renderContent, runs, paragraph, nativeFontMap);
                RenderNativeHeading(pdf, headingLevel, headingText, objectAlign, headingColor, paragraph, paragraphStyle, nativeFontMap, headingLink.LinkUri, headingLink.LinkDestinationName, headingLink.LinkContents);
                if (CreateNativeBottomBorderRuleStyle(paragraph, paragraphStyle) is { } headingRuleStyle) {
                    pdf.HR(style: headingRuleStyle);
                }

                RenderNativeFormFields(pdf, formFieldControls, objectAlign);
                RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
                RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
                return;
            }

            PdfCore.PanelStyle? panelStyle = CreateNativeParagraphPanelStyle(paragraph, paragraphStyle);
            if (panelStyle != null) {
                pdf.PanelParagraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, options, nativeDefaults, nativeFontMap);
                }, panelStyle, align, defaultColor);
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

            if (hasRenderableRuns || !string.IsNullOrEmpty(renderContent) || marker != null || paragraphFootnoteNumbers.Count > 0) {
                pdf.Paragraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, options, nativeDefaults, nativeFontMap);
                }, align, defaultColor, paragraphStyle);
            }

            if (bottomBorderRuleStyle != null) {
                pdf.HR(style: bottomBorderRuleStyle);
            }

            RenderNativeFormFields(pdf, formFieldControls, objectAlign);
            RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
            RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
        }

        private static void RenderNativeEmptyParagraph(INativePdfFlow pdf, WordParagraph paragraph, PdfCore.PdfParagraphStyle style, NativeDocumentDefaults nativeDefaults, bool renderSpacingOnlyLineBox) {
            if (!ShouldRenderNativeEmptyParagraphLineBox(paragraph, renderSpacingOnlyLineBox)) {
                if (renderSpacingOnlyLineBox && paragraph.LineSpacingAfterPoints is { } spacingAfter && spacingAfter > 0D) {
                    pdf.Spacer(spacingAfter);
                }

                return;
            }

            double height = MeasureNativeEmptyParagraphHeight(paragraph, style, nativeDefaults);
            if (height > 0D) {
                pdf.Spacer(height);
            }
        }

        private static bool ShouldRenderNativeEmptyParagraphLineBox(WordParagraph paragraph, bool renderSpacingOnlyLineBox) {
            if (paragraph.FontSize.HasValue ||
                paragraph.LineSpacingBeforePoints.HasValue ||
                paragraph.LineSpacingPoints.HasValue ||
                paragraph.LineSpacing.HasValue) {
                return true;
            }

            if (paragraph._paragraph != null &&
                paragraph._paragraph.ParagraphProperties == null &&
                !paragraph._paragraph.Elements<W.Run>().Any()) {
                return true;
            }

            return paragraph._paragraph?.ParagraphProperties?.ParagraphMarkRunProperties != null;
        }

        private static double MeasureNativeEmptyParagraphHeight(WordParagraph paragraph, PdfCore.PdfParagraphStyle style, NativeDocumentDefaults nativeDefaults) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults);
            double lineHeight = style.LineHeight ?? ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, styleDefaults);
            double spacingAfter = style.SpacingAfter ?? nativeDefaults.ParagraphSpacingAfter;
            double height = style.SpacingBefore + (fontSize * lineHeight) + spacingAfter;
            return double.IsNaN(height) || double.IsInfinity(height) ? 0D : Math.Max(0D, height);
        }

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

        private static IReadOnlyList<PdfCore.PdfTableCellCheckBox> CreateNativeTableCellCheckBoxes(WordTableCell cell) {
            var checkBoxes = new List<PdfCore.PdfTableCellCheckBox>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                IReadOnlyList<W.SdtRun> controls = GetNativeCheckBoxControls(paragraph);
                for (int index = 0; index < controls.Count; index++) {
                    W.SdtRun checkbox = controls[index];
                    checkBoxes.Add(new PdfCore.PdfTableCellCheckBox(
                        GetNativeCheckBoxFieldName(checkbox, checkBoxes.Count, "WordTableCheckBox"),
                        IsNativeCheckBoxChecked(checkbox),
                        size: 12D));
                }
            }

            return checkBoxes;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellFormField> CreateNativeTableCellFormFields(WordTableCell cell) {
            var formFields = new List<PdfCore.PdfTableCellFormField>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                IReadOnlyList<W.SdtRun> controls = GetNativeFormFieldControls(paragraph);
                for (int index = 0; index < controls.Count; index++) {
                    W.SdtRun formField = controls[index];
                    if (IsNativeDatePickerControl(formField)) {
                        formFields.Add(PdfCore.PdfTableCellFormField.TextField(
                            GetNativeContentControlFieldName(formField, formFields.Count, "WordTableDatePicker"),
                            GetNativeDatePickerValue(formField),
                            width: 150D,
                            height: 20D,
                            fontSize: 10D));
                        continue;
                    }

                    IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                    string? value = GetNativeChoiceFieldValue(formField, options);
                    if (options.Count == 0 || string.IsNullOrWhiteSpace(value)) {
                        continue;
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
            }

            return formFields;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellImage> CreateNativeTableCellImages(WordTableCell cell) {
            var images = new List<PdfCore.PdfTableCellImage>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                if (paragraph.Image != null) {
                    AddNativeTableCellImage(images, paragraph.Image);
                }

                foreach (W.SdtRun pictureControl in GetNativePictureControls(paragraph)) {
                    var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, pictureControl);
                    WordImage? pictureControlImage = pictureParagraph.PictureControl?.Image;
                    if (pictureControlImage == null) {
                        continue;
                    }

                    AddNativeTableCellImage(images, pictureControlImage);
                }
            }

            return images;
        }

        private static void AddNativeTableCellImage(List<PdfCore.PdfTableCellImage> images, WordImage image) {
            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!IsNativePdfSupportedImageBytes(bytes, out _)) {
                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            images.Add(new PdfCore.PdfTableCellImage(bytes, width, height, CreateNativeImageStyle()));
        }

        private static void AddNativeParagraphContent(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph paragraph,
            (int Level, string Marker)? marker,
            IReadOnlyList<WordParagraph> runs,
            bool hasRenderableRuns,
            string content,
            IReadOnlyList<int> paragraphFootnoteNumbers,
            PdfSaveOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            if (marker != null) {
                builder.Text(new string(' ', Math.Max(0, marker.Value.Level - 1) * 2));
                builder.Text(marker.Value.Marker);
                builder.Text(" ");
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
                        builder.LineBreak();
                        tabIndex = 0;
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

            AddNativeFootnoteReferences(builder, paragraphFootnoteNumbers);
        }

        private static void RenderNativeRunImages(INativePdfFlow pdf, IReadOnlyList<WordParagraph> runs, PdfCore.PdfAlign align, PdfSaveOptions? options) {
            foreach (WordParagraph run in runs) {
                if (run.IsImage && run.Image != null) {
                    RenderNativeImage(pdf, run.Image, align, options, "body paragraph image run");
                }
            }
        }

        private static void RenderNativeRunCharts(INativePdfFlow pdf, IReadOnlyList<WordParagraph> runs, PdfCore.PdfAlign align, PdfSaveOptions? options, W.Run? currentRun = null) {
            foreach (WordParagraph run in runs) {
                if (currentRun != null && ReferenceEquals(run._run, currentRun)) {
                    continue;
                }

                RenderNativeChart(pdf, run.Chart, align, options, "body paragraph chart run");
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

        private static void AddNativeFootnoteReferences(PdfCore.PdfParagraphBuilder builder, IReadOnlyList<int> footnoteNumbers) {
            foreach (int footnoteNumber in footnoteNumbers) {
                builder.Baseline(PdfCore.PdfTextBaseline.Superscript);
                builder.Text(footnoteNumber.ToString(CultureInfo.InvariantCulture));
                builder.Baseline(PdfCore.PdfTextBaseline.Normal);
            }
        }

        private static bool IsNativeTextWrappingBreak(WordParagraph run) =>
            run.IsBreak && run.Break?.BreakType != W.BreakValues.Page;

        private static WordTextBox? GetNativeParagraphTextBox(WordParagraph paragraph, out string? fallbackText) {
            fallbackText = GetNativeParagraphTextBoxPlainText(paragraph);
            WordTextBox? textBox = paragraph.TextBox;
            if (textBox != null || paragraph._paragraph == null) {
                return textBox;
            }

            foreach (W.Run run in paragraph._paragraph.Elements<W.Run>()) {
                if (run.Descendants<Wps.TextBoxInfo2>().Any() ||
                    run.Descendants<DocumentFormat.OpenXml.Vml.TextBox>().Any()) {
                    return new WordTextBox(paragraph._document, paragraph._paragraph, run);
                }
            }

            return null;
        }

        private static string? GetNativeParagraphTextBoxPlainText(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return null;
            }

            var parts = new List<string>();
            foreach (Wps.TextBoxInfo2 textBoxInfo in paragraph._paragraph.Descendants<Wps.TextBoxInfo2>()) {
                parts.AddRange(textBoxInfo.Descendants<W.Text>().Select(text => text.Text));
            }

            foreach (DocumentFormat.OpenXml.Vml.TextBox textBox in paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Vml.TextBox>()) {
                parts.AddRange(textBox.Descendants<W.Text>().Select(text => text.Text));
            }

            string textBoxText = ResolveNativeBuiltInPropertyPlaceholders(paragraph._document, string.Concat(parts));
            return string.IsNullOrWhiteSpace(textBoxText) ? null : textBoxText;
        }

        private static void RenderNativeTextBox(INativePdfFlow pdf, WordTextBox textBox, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap, string? fallbackText = null) {
            if (!string.IsNullOrWhiteSpace(fallbackText)) {
                PdfCore.PanelStyle fallbackStyle = CreateNativeTextBoxPanelStyle(textBox);
                pdf.PanelParagraph(builder => builder.Text(NormalizeNativeDirectText(fallbackText)), fallbackStyle, PdfCore.PdfAlign.Left);
                return;
            }

            IReadOnlyList<WordParagraph> paragraphs = GetNativeTextBoxParagraphs(textBox);
            if (paragraphs.Count == 0) {
                return;
            }

            PdfCore.PanelStyle style = CreateNativeTextBoxPanelStyle(textBox);
            PdfCore.PdfAlign defaultTextAlign = MapNativeTextBoxTextAlign(paragraphs);
            pdf.PanelParagraph(builder => {
                for (int index = 0; index < paragraphs.Count; index++) {
                    WordParagraph paragraph = paragraphs[index];
                    if (index > 0) {
                        builder.LineBreak();
                    }

                    List<WordParagraph> runs = GetNativeRuns(paragraph);
                    bool hasEquationContent = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0;
                    string content = hasEquationContent
                        ? AppendNativeTextWithEquation(paragraph.Text, paragraph)
                        : paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                    bool hasRenderableRuns = runs.Any(run => IsNativeRenderableTextRun(run, paragraph));
                    string renderContent = hasRenderableRuns || ShouldRenderNativeDirectText(paragraph, runs, content) ? content : string.Empty;
                    List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                    AddNativeParagraphContent(builder, paragraph, null, runs, hasRenderableRuns, renderContent, paragraphFootnoteNumbers, options, nativeDefaults, nativeFontMap);
                }
            }, style, defaultTextAlign);
        }

        private static IReadOnlyList<WordParagraph> GetNativeTextBoxParagraphs(WordTextBox textBox) {
            IReadOnlyList<WordParagraph> directParagraphs = textBox.Paragraphs;
            if (HasNativeRenderableTextBoxText(directParagraphs)) {
                return directParagraphs;
            }

            IReadOnlyList<WordParagraph> elementParagraphs = CollapseNativeParagraphElements(textBox.Elements)
                .OfType<WordParagraph>()
                .ToList();
            return elementParagraphs;
        }

        private static bool HasNativeRenderableTextBoxText(IEnumerable<WordParagraph> paragraphs) {
            foreach (WordParagraph paragraph in paragraphs) {
                List<WordParagraph> runs = GetNativeRuns(paragraph);
                if (runs.Count == 0 && !IsNativeHiddenTextRun(paragraph) && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                    return true;
                }

                if (runs.Any(run => IsNativeRenderableTextRun(run, paragraph) && !string.IsNullOrWhiteSpace(run.Text))) {
                    return true;
                }
            }

            return false;
        }

        private static PdfCore.PanelStyle CreateNativeTextBoxPanelStyle(WordTextBox textBox) {
            var style = new PdfCore.PanelStyle {
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

        private static PdfCore.PdfAlign MapNativeTextBoxBoxAlign(WordHorizontalAlignmentValues alignment) {
            switch (alignment) {
                case WordHorizontalAlignmentValues.Center:
                    return PdfCore.PdfAlign.Center;
                case WordHorizontalAlignmentValues.Right:
                case WordHorizontalAlignmentValues.Outside:
                    return PdfCore.PdfAlign.Right;
                default:
                    return PdfCore.PdfAlign.Left;
            }
        }

        private static PdfCore.PdfAlign MapNativeTextBoxTextAlign(IReadOnlyList<WordParagraph> paragraphs) {
            foreach (WordParagraph paragraph in paragraphs) {
                if (!string.IsNullOrEmpty(paragraph.Text)) {
                    return ResolveNativeParagraphAlign(paragraph);
                }
            }

            return PdfCore.PdfAlign.Left;
        }

        private static void RenderNativeHeading(INativePdfFlow pdf, int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle, NativeFontMap nativeFontMap, string? linkUri = null, string? linkDestinationName = null, string? linkContents = null) {
            PdfCore.PdfHeadingStyle style = CreateNativeWordHeadingStyle(level, paragraph, paragraphStyle, nativeFontMap);
            string normalizedText = NormalizeNativeDirectText(text);
            if (string.IsNullOrWhiteSpace(normalizedText)) {
                return;
            }

            pdf.Heading(level, normalizedText, align, color, style, linkUri, linkDestinationName, linkContents);
        }

        private static string GetNativeHeadingText(string content, IReadOnlyList<WordParagraph> runs, WordParagraph paragraph, NativeFontMap nativeFontMap) {
            string normalizedContent = NormalizeNativeDirectText(content);
            if (!string.IsNullOrWhiteSpace(normalizedContent)) {
                return ApplyNativeTextTransform(normalizedContent, paragraph, nativeFontMap: nativeFontMap);
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

            return NormalizeNativeDirectText(builder.ToString());
        }

        private static bool IsNativeRenderableTextRun(WordParagraph run, WordParagraph? fallback = null) =>
            !run.IsImage &&
            !string.IsNullOrEmpty(run.Text) &&
            !IsNativeHiddenTextRun(run, fallback);

        private static bool ShouldRenderNativeDirectText(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, string content) =>
            runs.Count == 0 &&
            !string.IsNullOrEmpty(content) &&
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

        private static PdfCore.PdfHeadingStyle CreateNativeWordHeadingStyle(int level, WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle, NativeFontMap nativeFontMap) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = level switch {
                1 => 16D,
                2 => 13D,
                _ => 12D
            };

            var style = new PdfCore.PdfHeadingStyle {
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

            if (HasNativeHeadingDeclaredLineHeight(paragraph, styleDefaults) && paragraphStyle.LineHeight.HasValue) {
                style.LineHeight = paragraphStyle.LineHeight.Value;
            }

            if (HasNativeHeadingDeclaredSpacingBefore(paragraph, styleDefaults)) {
                style.SpacingBefore = paragraphStyle.SpacingBefore;
            }

            if (HasNativeHeadingDeclaredSpacingAfter(paragraph, styleDefaults)) {
                style.SpacingAfter = paragraphStyle.SpacingAfter;
            }

            style.KeepWithNext = ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? styleDefaults.KeepWithNext ?? true;
            string? headingFontFamily = ResolveNativeParagraphStyleFontFamily(paragraph._document, paragraph.StyleId);
            if (nativeFontMap.TryGetFontSlot(headingFontFamily, out PdfCore.PdfStandardFont headingFont)) {
                style.Font = headingFont;
            }

            return style;
        }

        private static double? ResolveNativeHeadingDeclaredFontSize(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) {
            if (paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0D) {
                return paragraph.FontSize.Value;
            }

            if (styleDefaults.FontSize.HasValue && styleDefaults.FontSize.Value > 0D) {
                return styleDefaults.FontSize.Value;
            }

            return null;
        }

        private static bool HasNativeHeadingDeclaredLineHeight(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.LineSpacing.HasValue ||
            paragraph.LineSpacingPoints.HasValue ||
            styleDefaults.LineHeight.HasValue ||
            styleDefaults.LineSpacingPoints.HasValue;

        private static bool HasNativeHeadingDeclaredSpacingBefore(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.LineSpacingBeforePoints.HasValue || styleDefaults.SpacingBefore.HasValue;

        private static bool HasNativeHeadingDeclaredSpacingAfter(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.LineSpacingAfterPoints.HasValue || styleDefaults.SpacingAfter.HasValue;

        private static (string? LinkUri, string? LinkDestinationName, string? LinkContents) GetNativeHeadingLink(WordParagraph paragraph) {
            if (!paragraph.IsHyperLink || paragraph.Hyperlink == null) {
                return (null, null, null);
            }

            string? contents = string.IsNullOrWhiteSpace(paragraph.Hyperlink.Tooltip)
                ? paragraph.Hyperlink.Text
                : paragraph.Hyperlink.Tooltip;
            if (string.IsNullOrWhiteSpace(contents)) {
                contents = null;
            }

            Uri? uri = paragraph.Hyperlink.Uri;
            if (uri != null && uri.IsAbsoluteUri) {
                return (uri.AbsoluteUri, null, contents);
            }

            string? bookmarkName = paragraph.Hyperlink.Anchor;
            if (!string.IsNullOrWhiteSpace(bookmarkName)) {
                return (null, bookmarkName, contents);
            }

            return (null, null, null);
        }

        private static void AddNativeRun(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph run,
            WordParagraph paragraphStyleFallback,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            PdfSaveOptions? options,
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
            PdfSaveOptions? options,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            if (string.IsNullOrEmpty(text) || IsNativeHiddenTextRun(run, paragraphStyleFallback)) {
                return;
            }

            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraphStyleFallback, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
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
            PdfSaveOptions? options,
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

        private readonly record struct NativeResolvedTextStyle(
            bool Bold,
            bool Underline,
            bool Italic,
            bool Strike,
            bool AllCaps,
            PdfCore.PdfTextBaseline Baseline,
            double? FontSize,
            PdfCore.PdfStandardFont? Font,
            PdfCore.PdfColor? Color,
            PdfCore.PdfColor? BackgroundColor);

        private static void AddNativeText(
            PdfCore.PdfParagraphBuilder builder,
            string text,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            ApplyNativeTextStyle(builder, paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            AddNativeRunText(builder, ApplyNativeTextTransform(text, paragraph, nativeFontMap: nativeFontMap), tabStops, ref tabIndex);
            ResetNativeTextStyle(builder);
        }

        private static string ApplyNativeTextTransform(string text, WordParagraph paragraph, WordParagraph? fallback = null, NativeTableRunStyleDefaults tableRunStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) =>
            ResolveNativeTextRunStyle(paragraph, fallback, tableRunStyleDefaults, nativeDefaults, nativeFontMap).AllCaps
                ? text.ToUpperInvariant()
                : text;

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, WordParagraph paragraph, WordParagraph? fallback = null, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) =>
            ApplyNativeTextStyle(builder, ResolveNativeTextRunStyle(paragraph, fallback, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap));

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, NativeResolvedTextStyle style) {
            builder.Bold(style.Bold);
            builder.Italic(style.Italic);
            builder.Underline(style.Underline);
            builder.Strike(style.Strike);
            builder.Baseline(style.Baseline);
            if (style.FontSize.HasValue) {
                builder.FontSize(style.FontSize.Value);
            }

            if (style.Font.HasValue) {
                builder.Font(style.Font.Value);
            }

            if (style.Color.HasValue) {
                builder.Color(style.Color.Value);
            }

            if (style.BackgroundColor.HasValue) {
                builder.BackgroundColor(style.BackgroundColor.Value);
            }
        }

        private static NativeResolvedTextStyle ResolveNativeTextRunStyle(WordParagraph paragraph, WordParagraph? fallback = null, NativeTableRunStyleDefaults tableRunStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null, NativeFontMap? nativeFontMap = null) {
            WordParagraph styleSource = fallback ?? paragraph;
            NativeDocumentDefaults resolvedNativeDefaults = nativeDefaults ?? GetNativeDocumentDefaults(styleSource._document);
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(styleSource);
            W.RunProperties? runProperties = GetNativeRunProperties(paragraph);
            NativeCharacterStyleDefaults characterStyleDefaults = GetNativeCharacterStyleDefaults(paragraph._document, runProperties);

            bool bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>()) ?? characterStyleDefaults.Bold ?? styleDefaults.Bold ?? tableRunStyleDefaults.Bold ?? false;
            bool italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>()) ?? characterStyleDefaults.Italic ?? styleDefaults.Italic ?? tableRunStyleDefaults.Italic ?? false;
            bool underline = ReadNativeUnderline(runProperties?.GetFirstChild<W.Underline>()) ?? characterStyleDefaults.Underline ?? styleDefaults.Underline ?? tableRunStyleDefaults.Underline ?? false;
            bool strike =
                ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>()) ??
                ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>()) ??
                characterStyleDefaults.Strike ??
                styleDefaults.Strike ??
                tableRunStyleDefaults.Strike ??
                false;
            bool allCaps =
                ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ??
                ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>()) ??
                characterStyleDefaults.AllCaps ??
                styleDefaults.AllCaps ??
                tableRunStyleDefaults.AllCaps ??
                false;
            PdfCore.PdfTextBaseline baseline = MapNativeTextBaseline(
                runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ??
                characterStyleDefaults.Baseline ??
                styleDefaults.Baseline ??
                tableRunStyleDefaults.Baseline);
            double? fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0
                ? paragraph.FontSize.Value
                : characterStyleDefaults.FontSize ?? styleDefaults.FontSize ?? tableRunStyleDefaults.FontSize;
            PdfCore.PdfStandardFont? font = ResolveNativeTextRunFont(paragraph, fallback, characterStyleDefaults, styleDefaults, tableRunStyleDefaults, resolvedNativeDefaults, nativeFontMap);

            PdfCore.PdfColor? color = TryGetNativeRunColor(runProperties, out PdfCore.PdfColor? directColor)
                ? directColor
                : ParseNativeColor(characterStyleDefaults.ColorHex) ?? ParseNativeColor(styleDefaults.ColorHex) ?? tableRunStyleDefaults.Color ?? ParseNativeColor(tableRunStyleDefaults.ColorHex);
            PdfCore.PdfColor? background = TryGetNativeRunHighlight(runProperties, out PdfCore.PdfColor? directBackground)
                ? directBackground
                : MapNativeHighlight(characterStyleDefaults.Highlight) ?? MapNativeHighlight(styleDefaults.Highlight) ?? MapNativeHighlight(tableRunStyleDefaults.Highlight);

            return new NativeResolvedTextStyle(bold, underline, italic, strike, allCaps, baseline, fontSize, font, color, background);
        }

        private static PdfCore.PdfTextBaseline MapNativeTextBaseline(W.VerticalPositionValues? baseline) =>
            baseline == W.VerticalPositionValues.Superscript
                ? PdfCore.PdfTextBaseline.Superscript
                : baseline == W.VerticalPositionValues.Subscript
                    ? PdfCore.PdfTextBaseline.Subscript
                    : PdfCore.PdfTextBaseline.Normal;

        private static W.RunProperties? GetNativeRunProperties(WordParagraph paragraph) =>
            paragraph.IsHyperLink ? paragraph.Hyperlink?._runProperties : paragraph._runProperties;

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

        private static void ResetNativeTextStyle(PdfCore.PdfParagraphBuilder builder) {
            builder.Bold(false)
                .Italic(false)
                .Underline(false)
                .Strike(false)
                .Baseline(PdfCore.PdfTextBaseline.Normal)
                .ResetColor()
                .ResetFontSize()
                .ResetFont()
                .ResetBackgroundColor();
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

        private static PdfCore.TextRun CreateNativeHyperLinkTextRun(
            string text,
            string? linkUri,
            string? bookmarkName,
            string? contents,
            NativeResolvedTextStyle style) =>
            new PdfCore.TextRun(
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
                backgroundColor: style.BackgroundColor);

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
