using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;


namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool TryRenderNativeParagraphFlowBreaks(
        INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker,
        System.Func<WordParagraph, (int Level, string Marker)?> getMarker,
        Dictionary<long, int> footnoteNumbersById, WordToPdfOptions? options,
        IReadOnlyDictionary<W.Paragraph, string> headingDestinations,
        NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap,
        bool renderSpacingOnlyEmptyParagraphLineBox, WordParagraph? nextParagraph) {
        if (!TrySplitNativeParagraphAtVisibleFlowBreak(paragraph, out W.BreakValues breakType, out WordParagraph? before, out WordParagraph? after)) return false;

        if (HasNativePageBreakBefore(paragraph)) pdf.PageBreak();
        if (paragraph._paragraph != null && string.IsNullOrEmpty(paragraph.Bookmark?.Name) &&
            headingDestinations.TryGetValue(paragraph._paragraph, out string? generatedDestination)) pdf.Bookmark(generatedDestination);
        while (true) {
            if (before != null) {
                before.PageBreakBeforeOverride = false;
                before.KeepWithNextOverride = false;
                before.LineSpacingAfterPoints = 0;
                RenderNativeParagraph(pdf, before, marker, getMarker, System.Array.Empty<int>(), footnoteNumbersById,
                    options, headingDestinations, nativeDefaults, nativeFontMap, renderSpacingOnlyEmptyParagraphLineBox, null);
                marker = null;
            }
            if (breakType == W.BreakValues.Column) pdf.ColumnBreak();
            else pdf.PageBreak(preserveEmptyPage: true);
            if (after == null) return true;
            after.PageBreakBeforeOverride = false;
            after.LineSpacingBeforePoints = 0;
            W.ParagraphProperties properties = after._paragraph!.ParagraphProperties ??= new W.ParagraphProperties();
            W.Indentation indentation = properties.GetFirstChild<W.Indentation>() ?? new W.Indentation();
            if (indentation.Parent == null) properties.AddChild(indentation, true);
            indentation.FirstLine = "0";
            indentation.FirstLineChars = null;
            indentation.Hanging = null;
            indentation.HangingChars = null;
            if (!TrySplitNativeParagraphAtVisibleFlowBreak(after, out breakType, out before, out WordParagraph? remaining)) {
                RenderNativeParagraph(pdf, after, marker, getMarker, System.Array.Empty<int>(), footnoteNumbersById,
                    options, headingDestinations, nativeDefaults, nativeFontMap, false, nextParagraph);
                return true;
            }
            after = remaining;
            renderSpacingOnlyEmptyParagraphLineBox = false;
        }
    }

    private static bool TrySplitNativeParagraphAtVisibleFlowBreak(WordParagraph paragraph, out W.BreakValues breakType, out WordParagraph? before, out WordParagraph? after) {
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> children = run._visibleRunSourceChildren ?? run._run!.ChildElements;
            W.Break? boundary = children.OfType<W.Break>().FirstOrDefault(item => item.Type?.Value == W.BreakValues.Page || item.Type?.Value == W.BreakValues.Column);
            if (boundary?.Type?.Value is W.BreakValues type) {
                breakType = type;
                return TrySplitNativeParagraphAtVisibleBreak(paragraph, type, out before, out after);
            }
        }
        breakType = W.BreakValues.Page; before = null; after = null;
        return false;
    }

    private static bool TrySplitNativeParagraphAtVisibleBreak(WordParagraph paragraph, W.BreakValues breakType, out WordParagraph? before, out WordParagraph? after) {
        before = null;
        after = null;
        if (paragraph._paragraph?.Descendants<W.Break>().Any(wordBreak => wordBreak.Type?.Value == breakType) != true) return false;
        var visibleBreaks = new HashSet<W.Break>();
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> visibleChildren = run._visibleRunSourceChildren ?? run._run!.ChildElements;
            foreach (W.Break wordBreak in visibleChildren.OfType<W.Break>()) {
                if (wordBreak.Type?.Value == breakType) visibleBreaks.Add(wordBreak);
            }
        }
        if (visibleBreaks.Count == 0) { before = null; after = null; return false; }
        return TrySplitNativeParagraphAtBreak(paragraph, breakType, visibleBreaks, out before, out after);
    }
    private static bool TrySplitNativeParagraphAtBreak(WordParagraph paragraph, W.BreakValues breakType,
        HashSet<W.Break>? visibleBreaks, out WordParagraph? before, out WordParagraph? after) {
        before = null;
        after = null;
        if (paragraph._paragraph == null) {
            return false;
        }

        var beforeParagraph = new W.Paragraph();
        var afterParagraph = new W.Paragraph();
        if (paragraph._paragraph.ParagraphProperties != null) {
            beforeParagraph.Append((W.ParagraphProperties)paragraph._paragraph.ParagraphProperties.CloneNode(true));
            afterParagraph.Append((W.ParagraphProperties)paragraph._paragraph.ParagraphProperties.CloneNode(true));
        }

        bool sawBreak = false;
        foreach (DocumentFormat.OpenXml.OpenXmlElement child in paragraph._paragraph.ChildElements) {
            if (child is W.ParagraphProperties) {
                continue;
            }

            if (!sawBreak &&
                TrySplitNativeOpenXmlAtBreak(child, breakType, visibleBreaks, out DocumentFormat.OpenXml.OpenXmlElement? beforeChild, out DocumentFormat.OpenXml.OpenXmlElement? afterChild)) {
                if (beforeChild != null) {
                    beforeParagraph.Append(beforeChild);
                }

                if (afterChild != null) {
                    afterParagraph.Append(afterChild);
                }

                sawBreak = true;
                continue;
            }

            if (sawBreak) {
                afterParagraph.Append(child.CloneNode(true));
            } else {
                beforeParagraph.Append(child.CloneNode(true));
            }
        }

        if (!sawBreak) {
            return false;
        }

        W.Break boundary = paragraph._paragraph.Descendants<W.Break>().First(wordBreak =>
            wordBreak.Type?.Value == breakType && (visibleBreaks == null || visibleBreaks.Contains(wordBreak)));
        WordComplexFieldRunVisibility.PreserveFragmentContext(paragraph._paragraph, beforeParagraph, afterParagraph, boundary);

        if (HasNativeRenderableOpenXmlContent(beforeParagraph)) {
            before = new WordParagraph(paragraph._document, beforeParagraph) {
                _bookmarkStart = beforeParagraph.Elements<W.BookmarkStart>().FirstOrDefault()
            };
        }

        if (HasNativeRenderableOpenXmlContent(afterParagraph)) {
            after = new WordParagraph(paragraph._document, afterParagraph) {
                _bookmarkStart = afterParagraph.Elements<W.BookmarkStart>().FirstOrDefault()
            };
        }

        return true;
    }

    private static bool TrySplitNativeOpenXmlAtBreak(
        DocumentFormat.OpenXml.OpenXmlElement element,
        W.BreakValues breakType,
        HashSet<W.Break>? visibleBreaks,
        out DocumentFormat.OpenXml.OpenXmlElement? before,
        out DocumentFormat.OpenXml.OpenXmlElement? after) {
        before = null;
        after = null;
        if (element is W.Break wordBreak && wordBreak.Type?.Value == breakType &&
            (visibleBreaks == null || visibleBreaks.Contains(wordBreak))) {
            return true;
        }

        if (!element.HasChildren) {
            return false;
        }

        DocumentFormat.OpenXml.OpenXmlElement beforeElement = element.CloneNode(false);
        DocumentFormat.OpenXml.OpenXmlElement afterElement = element.CloneNode(false);
        bool sawBreak = false;
        bool containsBreak = false;
        foreach (DocumentFormat.OpenXml.OpenXmlElement child in element.ChildElements) {
            // Run/control properties apply to both fragments of their container.
            if (child is W.RunProperties or W.SdtProperties or W.SdtEndCharProperties) {
                beforeElement.Append(child.CloneNode(true));
                afterElement.Append(child.CloneNode(true));
                continue;
            }
            if (!sawBreak &&
                TrySplitNativeOpenXmlAtBreak(child, breakType, visibleBreaks, out DocumentFormat.OpenXml.OpenXmlElement? beforeChild, out DocumentFormat.OpenXml.OpenXmlElement? afterChild)) {
                containsBreak = true;
                if (beforeChild != null) {
                    beforeElement.Append(beforeChild);
                }

                if (afterChild != null) {
                    afterElement.Append(afterChild);
                }

                sawBreak = true;
                continue;
            }

            if (sawBreak) {
                afterElement.Append(child.CloneNode(true));
            } else {
                beforeElement.Append(child.CloneNode(true));
            }
        }

        if (!containsBreak) {
            return false;
        }

        if (HasNativeRenderableOpenXmlContent(beforeElement)) {
            before = beforeElement;
        }

        if (HasNativeRenderableOpenXmlContent(afterElement)) {
            after = afterElement;
        }

        return true;
    }

    private static bool HasNativeRenderableOpenXmlContent(DocumentFormat.OpenXml.OpenXmlElement element) {
        if (element.Descendants<W.Text>().Any(text => !string.IsNullOrEmpty(text.Text)) ||
            element.Descendants<W.BookmarkStart>().Any() ||
            element.Descendants<W.TabChar>().Any() ||
            element.Descendants<W.Drawing>().Any() ||
            element.Descendants<W.FootnoteReference>().Any() ||
            element.Descendants<W.EndnoteReference>().Any() ||
            element.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() ||
            element.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any()) {
            return true;
        }

        return element.Descendants<W.Break>().Any();
    }

}
