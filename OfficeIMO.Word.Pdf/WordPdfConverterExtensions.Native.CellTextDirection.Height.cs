using System.Collections.Generic;
using System.Globalization;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Word's automatic turned rows use the first paragraph mark's horizontal line box, rather than its turned visible runs.</summary>
    private static double? GetNativeOrientedCellMarkHeight(WordTableCell cell, IReadOnlyList<PdfCore.PdfTextRun> runs,
        NativeTableCellEmbeddedContent embedded, NativeDocumentDefaults defaults, NativeTableStyleDefaults tableStyle,
        NativeFontMap fontMap) {
        if (GetNativeCellTextRotation(cell.TextDirection) == 0 || cell.HideMark == true ||
            runs.Any(run => run.InlineElement != null) || embedded.Images.Count > 0 ||
            embedded.CheckBoxes.Count > 0 || embedded.FormFields.Count > 0) return null;
        WordParagraph? paragraph = cell.Paragraphs.FirstOrDefault();
        if (paragraph == null) return null;
        NativeParagraphStyleDefaults style = GetNativeParagraphStyleDefaults(paragraph);
        W.ParagraphMarkRunProperties? mark = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
        W.RunProperties? markRun = mark == null ? null : new W.RunProperties(mark.ChildElements.Select(child => child.CloneNode(true)));
        NativeCharacterStyleDefaults characterStyle = GetNativeCharacterStyleDefaults(paragraph._document, markRun);
        string? authoredSize = mark?.GetFirstChild<W.FontSize>()?.Val?.Value;
        double? sourceSize = double.TryParse(authoredSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
            ? halfPoints / 2D : characterStyle.FontSize ?? style.FontSize ?? tableStyle.RunStyle.FontSize;
        if (tableStyle.UseConfiguredTypography && !sourceSize.HasValue) return null;
        double size = sourceSize ?? defaults.FontSize;
        double natural = ResolveNativeWordSingleLineHeight(fontMap,
            EnumerateNativeLatinFontFamilies(paragraph._document, markRun?.RunFonts)
                .Concat(EnumerateNativeStyleFontFamilies(characterStyle, style, tableStyle.RunStyle, defaults,
                    fontMap.UsePdfDefaultForDocumentDefaultFont != true)).ToArray());
        double multiplier = ResolveNativeParagraphLineSpacing(paragraph, style, defaults, tableStyle.LineSpacing)
            .Resolve(size, natural) ?? natural;
        return size * multiplier;
    }
}
