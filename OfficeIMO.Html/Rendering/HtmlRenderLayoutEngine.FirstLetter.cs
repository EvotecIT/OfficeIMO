using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ApplyFirstLetterStyle(
        IElement? formattingContainer,
        double width,
        HtmlRenderBoxStyle parentStyle,
        List<HtmlInlineRun> runs) {
        if (formattingContainer == null
            || !TryResolveFirstLetterStyle(formattingContainer, width, parentStyle, out _)) return;

        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            if (IsInlineTextTransformContextMarker(run)) continue;
            if (run.AtomicBlock != null) return;
            if (run.Text.Length == 0) continue;
            if (!TryFindFirstLetterBounds(run.Text, run.Style.PreserveWhitespace, out int start, out int end,
                    out bool firstLineBlocked)) {
                if (firstLineBlocked) return;
                continue;
            }
            string prefix = run.Text.Substring(0, start);
            string firstLetter = run.Text.Substring(start, end - start);
            string suffix = run.Text.Substring(end);
            var replacement = new List<HtmlInlineRun>(3);
            if (prefix.Length > 0) replacement.Add(run.CloneText(prefix, prefix, run.Style));
            if (!TryResolveFirstLetterStyle(formattingContainer, width, run.Style, out HtmlRenderBoxStyle letterStyle)) return;
            if (run.IsTextTransformContextOnly && letterStyle.Font.Size > 0D) {
                ReportUnsupportedComplexTextShaping(firstLetter, run.OwnerElement ?? formattingContainer, letterStyle, run.Source);
            }
            replacement.Add(run.CloneText(firstLetter, firstLetter, letterStyle, isFirstLetter: true));
            if (suffix.Length > 0) replacement.Add(run.CloneText(suffix, suffix, run.Style));
            runs.RemoveAt(runIndex);
            runs.InsertRange(runIndex, replacement);
            return;
        }
    }

    private bool TryFindFirstLetterBounds(string text, bool preserveWhitespace, out int start, out int end,
        out bool firstLineBlocked) {
        start = end = 0;
        firstLineBlocked = false;
        TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(text);
        int examined = 0;
        bool MoveNext() {
            if (!enumerator.MoveNext()) return false;
            ChargeLayoutOperation("first-letter scanning");
            if ((++examined & 0xFF) == 0) CheckCancellation();
            return true;
        }

        if (!MoveNext()) return false;
        while (string.IsNullOrWhiteSpace(enumerator.GetTextElement())) {
            string whitespace = enumerator.GetTextElement();
            if (whitespace.IndexOf('\u2028') >= 0
                || preserveWhitespace && (whitespace.IndexOf('\n') >= 0 || whitespace.IndexOf('\r') >= 0)) {
                firstLineBlocked = true;
                return false;
            }
            if (!MoveNext()) return false;
        }
        start = enumerator.ElementIndex;
        while (IsFirstLetterPunctuation(enumerator.GetTextElement())) {
            if (!MoveNext()) {
                end = text.Length;
                return true;
            }
        }
        if (!MoveNext()) {
            end = text.Length;
            return true;
        }
        while (IsFirstLetterPunctuation(enumerator.GetTextElement())) {
            if (!MoveNext()) {
                end = text.Length;
                return true;
            }
        }
        end = enumerator.ElementIndex;
        return end > start;
    }

    // The fictional first-letter box abuts the selected text, including text in
    // an inline descendant. Its inherited font must therefore come from that run.
    private bool TryResolveFirstLetterStyle(
        IElement formattingContainer,
        double width,
        HtmlRenderBoxStyle inheritedStyle,
        out HtmlRenderBoxStyle letterStyle) {
        if (!_styleResolver.TryResolvePseudo(formattingContainer, HtmlPseudoElementKind.FirstLetter,
                width, inheritedStyle, out letterStyle)) return false;
        letterStyle.Language = inheritedStyle.Language;
        return true;
    }

    private static bool IsFirstLetterPunctuation(string textElement) {
        if (string.IsNullOrEmpty(textElement)) return false;
        UnicodeCategory category = char.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.OpenPunctuation
            || category == UnicodeCategory.ClosePunctuation
            || category == UnicodeCategory.InitialQuotePunctuation
            || category == UnicodeCategory.FinalQuotePunctuation
            || category == UnicodeCategory.OtherPunctuation;
    }
}
