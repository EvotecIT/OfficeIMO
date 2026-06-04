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
        private static void RenderNativeElement(INativePdfFlow pdf, WordElement element, Func<WordParagraph, (int Level, string Marker)?> getMarker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            switch (element) {
                case WordParagraph paragraph:
                    RenderNativeParagraph(pdf, paragraph, getMarker(paragraph), footnoteNumbers, footnoteNumbersById, options, headingDestinations);
                    break;
                case WordTableOfContent tableOfContent:
                    RenderNativeTableOfContents(pdf, tableOfContent, tableOfContentsEntries);
                    break;
                case WordTable table:
                    RenderNativeTable(pdf, table, getMarker, footnoteNumbersById, options);
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

        private static void RenderNativeParagraph(INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            if (paragraph == null) {
                return;
            }

            if (paragraph.PageBreakBefore) {
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
                RenderNativeTextBox(pdf, textBox, footnoteNumbersById, options, textBoxFallbackText);
                return;
            }

            if (paragraph.Shape != null) {
                RenderNativeShape(pdf, paragraph.Shape);
            }

            if (paragraph.Image != null) {
                RenderNativeImage(pdf, paragraph.Image, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options, "body paragraph image");
            }

            WordImage? pictureControlImage = paragraph.PictureControl?.Image;
            if (pictureControlImage != null) {
                RenderNativeImage(pdf, pictureControlImage, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options, "body picture control image");
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (paragraph.Image == null) {
                RenderNativeRunImages(pdf, runs, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options);
            }

            string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : AppendNativeTextWithEquation(paragraph.Text, paragraph);
            bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
            List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, footnoteNumbers, footnoteNumbersById);
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph);
            if (marker == null &&
                paragraphFootnoteNumbers.Count == 0 &&
                IsNativeHorizontalRuleParagraph(paragraph, runs, content) &&
                CreateNativeHorizontalRuleStyle(paragraph, style) is { } horizontalRuleStyle) {
                pdf.HR(style: horizontalRuleStyle);
                return;
            }

            if (!hasRenderableRuns && string.IsNullOrEmpty(content) && marker == null && paragraphFootnoteNumbers.Count == 0 && checkboxControls.Count == 0 && formFieldControls.Count == 0 && repeatingSectionControls.Count == 0) {
                return;
            }

            PdfCore.PdfAlign align = MapNativeParagraphAlign(paragraph.ParagraphAlignment);
            PdfCore.PdfAlign objectAlign = MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false);
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

                RenderNativeHeading(pdf, headingLevel, content, objectAlign, headingColor, headingLink.LinkUri, headingLink.LinkDestinationName, headingLink.LinkContents);
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
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
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

            if (hasRenderableRuns || !string.IsNullOrEmpty(content) || marker != null || paragraphFootnoteNumbers.Count > 0) {
                pdf.Paragraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
                }, align, defaultColor, paragraphStyle);
            }

            if (bottomBorderRuleStyle != null) {
                pdf.HR(style: bottomBorderRuleStyle);
            }

            RenderNativeFormFields(pdf, formFieldControls, objectAlign);
            RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
            RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
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

                    pdf.Paragraph(builder => builder.Text(itemText), align, color);
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
            PdfSaveOptions? options) {
                if (marker != null) {
                    builder.Text(new string(' ', Math.Max(0, marker.Value.Level - 1) * 2));
                    builder.Text(marker.Value.Marker);
                    builder.Text(" ");
                }

                IReadOnlyList<WordTabStop> tabStops = paragraph.TabStops;
                int tabIndex = 0;
                if (hasRenderableRuns) {
                    foreach (WordParagraph run in runs) {
                        if (run.IsImage && run.Image != null) {
                            continue;
                        }

                        if (IsNativeTextWrappingBreak(run)) {
                            builder.LineBreak();
                            tabIndex = 0;
                            continue;
                        }

                        AddNativeRun(builder, run, paragraph, tabStops, ref tabIndex, options);
                }
                    string? supplementalText = GetNativeSupplementalTextAfterRuns(content, runs);
                    if (!string.IsNullOrEmpty(supplementalText)) {
                        AddNativeText(builder, supplementalText!, paragraph, tabStops, ref tabIndex);
                    }
            } else if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                    ApplyNativeTextStyle(builder, paragraph);
                    AddNativeHyperLinkRun(builder, paragraph.Hyperlink.Text, paragraph.Hyperlink, tabStops, ref tabIndex);
                    ResetNativeTextStyle(builder);
                } else {
                    AddNativeText(builder, content, paragraph, tabStops, ref tabIndex);
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

        private static string? GetNativeSupplementalTextAfterRuns(string content, IReadOnlyList<WordParagraph> runs) {
            if (string.IsNullOrEmpty(content)) {
                return null;
            }

            var renderedText = new StringBuilder();
            foreach (WordParagraph run in runs) {
                if (run.IsImage || IsNativeTextWrappingBreak(run) || string.IsNullOrEmpty(run.Text)) {
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

            string textBoxText = string.Concat(parts);
            return string.IsNullOrWhiteSpace(textBoxText) ? null : textBoxText;
        }

        private static void RenderNativeTextBox(INativePdfFlow pdf, WordTextBox textBox, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, string? fallbackText = null) {
            if (!string.IsNullOrWhiteSpace(fallbackText)) {
                PdfCore.PanelStyle fallbackStyle = CreateNativeTextBoxPanelStyle(textBox);
                pdf.PanelParagraph(builder => builder.Text(fallbackText!), fallbackStyle, PdfCore.PdfAlign.Left);
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
                    string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                    bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
                    List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                    AddNativeParagraphContent(builder, paragraph, null, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
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
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    return true;
                }

                if (GetNativeRuns(paragraph).Any(run => !run.IsImage && !string.IsNullOrWhiteSpace(run.Text))) {
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
                    return MapNativeParagraphAlign(paragraph.ParagraphAlignment);
                }
            }

            return PdfCore.PdfAlign.Left;
        }

        private static void RenderNativeHeading(INativePdfFlow pdf, int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, string? linkUri = null, string? linkDestinationName = null, string? linkContents = null) {
            PdfCore.PdfHeadingStyle style = CreateNativeWordHeadingStyle(level);
            pdf.Heading(level, text, align, color, style, linkUri, linkDestinationName, linkContents);
        }

        private static PdfCore.PdfHeadingStyle CreateNativeWordHeadingStyle(int level) {
            double fontSize = level switch {
                1 => 16D,
                2 => 13D,
                _ => 12D
            };

            return new PdfCore.PdfHeadingStyle {
                FontSize = fontSize,
                LineHeight = 1.18D,
                SpacingBefore = level == 1 ? 24D : 10D,
                SpacingAfter = level == 1 ? 5D : 4D,
                Bold = false,
                ApplySpacingBeforeAtTop = true,
                KeepWithNext = true
            };
        }

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
            PdfSaveOptions? options) {
            if (string.IsNullOrEmpty(run.Text)) {
                return;
            }

            ApplyNativeTextStyle(builder, run, paragraphStyleFallback);

            if (run.IsHyperLink && run.Hyperlink != null) {
                AddNativeHyperLinkRun(builder, run.Text, run.Hyperlink, tabStops, ref tabIndex);
            } else {
                AddNativeRunText(builder, run.Text, tabStops, ref tabIndex);
            }

            ResetNativeTextStyle(builder);
        }

        private static void AddNativeText(
            PdfCore.PdfParagraphBuilder builder,
            string text,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex) {
            ApplyNativeTextStyle(builder, paragraph);
            AddNativeRunText(builder, text, tabStops, ref tabIndex);
            ResetNativeTextStyle(builder);
        }

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, WordParagraph paragraph, WordParagraph? fallback = null) {
            builder.Bold(paragraph.Bold);
            builder.Italic(paragraph.Italic);
            builder.Underline(paragraph.Underline != null);
            builder.Strike(paragraph.Strike || paragraph.DoubleStrike);
            builder.Baseline(GetNativeTextBaseline(paragraph));
            if (paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0) {
                builder.FontSize(paragraph.FontSize.Value);
            }

            if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(paragraph.FontFamily, out PdfCore.PdfStandardFont font) ||
                PdfCore.PdfStandardFontMapper.TryMapFontFamily(paragraph.FontFamilyHighAnsi, out font) ||
                PdfCore.PdfStandardFontMapper.TryMapFontFamily(paragraph.FontFamilyEastAsia, out font) ||
                PdfCore.PdfStandardFontMapper.TryMapFontFamily(paragraph.FontFamilyComplexScript, out font) ||
                (fallback != null && (
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(fallback.FontFamily, out font) ||
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(fallback.FontFamilyHighAnsi, out font) ||
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(fallback.FontFamilyEastAsia, out font) ||
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(fallback.FontFamilyComplexScript, out font)))) {
                builder.Font(font);
            }

            PdfCore.PdfColor? color = ParseNativeColor(paragraph.ColorHex);
            PdfCore.PdfColor? background = MapNativeHighlight(paragraph.Highlight);
            if (color.HasValue) {
                builder.Color(color.Value);
            }

            if (background.HasValue) {
                builder.BackgroundColor(background.Value);
            }
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

        private static void AddNativeHyperLinkRun(PdfCore.PdfParagraphBuilder builder, string text, WordHyperLink hyperlink, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            Uri? uri = hyperlink.Uri;
            string? linkUri = uri != null && uri.IsAbsoluteUri ? uri.AbsoluteUri : null;
            string? bookmarkName = string.IsNullOrWhiteSpace(hyperlink.Anchor) ? null : hyperlink.Anchor;
            if (linkUri == null && bookmarkName == null) {
                AddNativeRunText(builder, text, tabStops, ref tabIndex);
                return;
            }

            string? contents = GetNativeHyperLinkContents(hyperlink);
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => {
                    if (linkUri != null) {
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
