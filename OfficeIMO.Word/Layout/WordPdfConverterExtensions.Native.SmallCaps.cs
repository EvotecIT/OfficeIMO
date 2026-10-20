using System.Collections.Generic;
using System.Globalization;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private readonly record struct NativeCasedTextSegment(string Text, NativeResolvedTextStyle Style);

    private static IEnumerable<NativeCasedTextSegment> GetNativeCasedTextSegments(string text, NativeResolvedTextStyle style) {
        if (!style.SmallCaps || style.AllCaps || string.IsNullOrEmpty(text)) {
            yield return new NativeCasedTextSegment(style.AllCaps ? text.ToUpperInvariant() : text, style);
            yield break;
        }

        // Word synthesizes small capitals at 80% of the authored size, rounded
        // to a half point. Native Arial/Times controls at 9--30 points distinguish
        // this rule from both an exact scale and an all-capitals transformation.
        double smallSize = Math.Max(0.5D, Math.Round(style.EffectiveFontSize * 1.6D, MidpointRounding.AwayFromZero) / 2D);
        NativeResolvedTextStyle smallStyle = style with { FontSize = smallSize, EffectiveFontSize = smallSize };
        var buffer = new StringBuilder();
        bool? currentSmall = null;
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(text);
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            bool small = CharUnicodeInfo.GetUnicodeCategory(element, 0) == UnicodeCategory.LowercaseLetter ||
                char.IsWhiteSpace(element, 0);
            if (currentSmall.HasValue && currentSmall.Value != small) {
                yield return new NativeCasedTextSegment(buffer.ToString(), currentSmall.Value ? smallStyle : style);
                buffer.Clear();
            }
            currentSmall = small;
            buffer.Append(element.ToUpperInvariant());
        }
        if (buffer.Length > 0)
            yield return new NativeCasedTextSegment(buffer.ToString(), currentSmall == true ? smallStyle : style);
    }

    private static void AddNativeCasedText(PdfCore.PdfParagraphBuilder builder, string text, NativeResolvedTextStyle style,
        IReadOnlyList<WordTabStop> tabStops, ref int tabIndex, WordHyperLink? hyperlink = null) {
        foreach (NativeCasedTextSegment segment in GetNativeCasedTextSegments(text, style)) {
            ApplyNativeTextStyle(builder, segment.Style);
            if (hyperlink != null)
                AddNativeHyperLinkRun(builder, segment.Text, hyperlink, tabStops, ref tabIndex, segment.Style);
            else
                AddNativeRunText(builder, segment.Text, tabStops, ref tabIndex);
        }
    }
}
