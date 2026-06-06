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
        private const double NativeDefaultParagraphLineHeight = 1.15D;
        private const double NativeDefaultParagraphSpacingAfter = 8D;

        private interface INativePdfFlow {
            void PageBreak();
            void Bookmark(string name);
            void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null);
            void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null);
            void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null);
            void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents);
            void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style);
            void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style);
            void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style = null);
            void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style = null);
            void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName = "Yes", PdfCore.PdfFormFieldStyle? style = null);
            void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null);
            void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style);
            void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null);
        }

        private sealed class NativePdfDocumentFlow : INativePdfFlow {
            private readonly PdfCore.PdfDocument _pdf;

            public NativePdfDocumentFlow(PdfCore.PdfDocument pdf) {
                _pdf = pdf;
            }

            public void PageBreak() => _pdf.PageBreak();
            public void Bookmark(string name) => _pdf.Bookmark(name);
            public void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _pdf.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _pdf.Paragraph(build, align, defaultColor, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null) => _pdf.PanelParagraph(build, style, align, defaultColor);
            public void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents) {
                if (level == 1) _pdf.H1(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else if (level == 2) _pdf.H2(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else _pdf.H3(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style) => _pdf.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style) => _pdf.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _pdf.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _pdf.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _pdf.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _pdf.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _pdf.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _pdf.Image(bytes, width, height, align, style: CreateNativeImageStyle());
        }

        private sealed class NativePdfColumnFlow : INativePdfFlow {
            private readonly PdfCore.PdfRowColumnCompose _column;

            public NativePdfColumnFlow(PdfCore.PdfRowColumnCompose column) {
                _column = column;
            }

            public void PageBreak() => _column.PageBreak();
            public void Bookmark(string name) => _column.Bookmark(name);
            public void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _column.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _column.Paragraph(build, align, defaultColor, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null) => _column.PanelParagraph(build, style, align, defaultColor);
            public void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents) {
                if (level == 1) _column.H1(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else if (level == 2) _column.H2(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else _column.H3(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style) => _column.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style) => _column.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _column.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _column.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _column.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _column.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _column.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _column.Image(bytes, width, height, align, style: CreateNativeImageStyle());
        }

        private static PdfCore.PdfImageStyle CreateNativeImageStyle() => new() {
            ScaleDownToFit = true
        };

        private static PdfCore.PdfDocument CreateOfficeIMOPdfDocument(WordDocument document, PdfSaveOptions? options) {
            options?.ResetExportState();

            BuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(CreateNativeOptions(document, options))
                .Meta(
                    title: options?.Title ?? properties.Title,
                    author: options?.Author ?? properties.Creator,
                    subject: options?.Subject ?? properties.Subject,
                    keywords: BuildNativeKeywords(options, properties));

            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            Dictionary<WordParagraph, (int Level, int Index)> listIndices = DocumentTraversal.BuildListIndices(document);
            Dictionary<W.Paragraph, string> headingDestinations = BuildNativeHeadingDestinations(document);
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries = BuildNativeTableOfContentsEntries(document, options, headingDestinations);
            var footnoteNumbersById = new Dictionary<long, int>();
            foreach (WordSection section in document.Sections) {
                IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(section.Elements);
                List<PdfFootnote> footnotes = CollectNativeFootnotes(elements, footnoteNumbersById);
                pdf.Section(page => {
                    page.Size(GetNativePageSize(section, options));
                    page.Margin(GetNativeMargins(section, options));
                    ConfigureNativePageNumbering(page, section);
                    ConfigureNativeHeaderFooter(page, section, options);
                    var flow = new NativePdfDocumentFlow(pdf);

                    if (TryRenderNativeSectionColumns(
                        page,
                        section,
                        elements,
                        listMarkers,
                        listIndices,
                        footnoteNumbersById,
                        options,
                        tableOfContentsEntries,
                        headingDestinations)) {
                        RenderNativeFootnotes(flow, footnotes);
                        return;
                    }

                    for (int i = 0; i < elements.Count; i++) {
                        WordElement element = elements[i];
                        if (element is WordFootNote) {
                            continue;
                        }

                        if (TryRenderNativeList(
                            flow,
                            elements,
                            ref i,
                            listMarkers,
                            listIndices,
                            footnoteNumbersById)) {
                            continue;
                        }

                        RenderNativeElement(
                            flow,
                            element,
                            paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                            GetNativeFootnoteNumbersForElement(elements, i, footnoteNumbersById),
                            footnoteNumbersById,
                            options,
                            tableOfContentsEntries,
                            headingDestinations);
                    }

                    RenderNativeFootnotes(flow, footnotes);
                });
            }

            return pdf;
        }

        private static IReadOnlyList<WordElement> CollapseNativeParagraphElements(IEnumerable<WordElement> elements) {
            var collapsed = new List<WordElement>();
            var paragraphIndexes = new Dictionary<W.Paragraph, int>();

            foreach (WordElement element in elements) {
                if (element is WordParagraph paragraph && paragraph._paragraph != null) {
                    if (paragraphIndexes.TryGetValue(paragraph._paragraph, out int existingIndex)) {
                        if (ShouldReplaceNativeParagraphElement(collapsed[existingIndex], paragraph)) {
                            collapsed[existingIndex] = paragraph;
                        }

                        continue;
                    }

                    paragraphIndexes.Add(paragraph._paragraph, collapsed.Count);
                }

                collapsed.Add(element);
            }

            return collapsed;
        }

        private static bool ShouldReplaceNativeParagraphElement(WordElement existing, WordParagraph candidate) {
            if (existing is not WordParagraph existingParagraph) {
                return false;
            }

            if (string.IsNullOrEmpty(existingParagraph.Bookmark?.Name) &&
                !string.IsNullOrEmpty(candidate.Bookmark?.Name)) {
                return true;
            }

            return false;
        }

    }
}
