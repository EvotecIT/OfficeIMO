using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Layout retains U+200B as a break opportunity. Font operations see only
    // painted text; linguistic join controls remain part of that text.
    internal static string WithoutOptionalBreaks(string text) => text.Replace("\u200B", string.Empty);

    // A single unspaced identifier may wrap at authored zero-width opportunities.
    // Its selection/search text must remain the identifier, without those layout controls.
    // Use only the rendered slice, so a continued paragraph never repeats its whole source.
    private static bool TryBeginOptionalBreakText(StringBuilder sb, RichParagraphBlock block,
        List<List<RichSeg>> lines, bool suppressed, PdfCreateOptions options, double x, double y, double width, double height) {
        if (suppressed) return false;
        string source = string.Concat(block.Runs.Select(run => run.Text));
        if (source.IndexOf('\u200B') < 0 || source.Any(char.IsWhiteSpace) || block.Runs.Any(run => run.InlineElement != null)) return false;
        string logicalText = string.Concat(lines.SelectMany(line => line).Select(segment => segment.Text)).Replace("\u200B", string.Empty);
        if (logicalText.Length == 0) return false;
        sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(logicalText)).Append(" >> BDC\n");
        PdfStandardFont font = ChooseNormal(options.DefaultFont);
        string fontResource = GetFontResourceName(font, null, font);
        WriteLogicalTextAnchor(sb, options, font, fontResource, x, y, logicalText, width, height).RestoreState();
        return true;
    }
    private static void EndOptionalBreakText(StringBuilder sb, RichParagraphBlock block, List<List<RichSeg>> lines,
        PdfCreateOptions options, double x, double y, double width, double height) {
        string logicalText = string.Concat(lines.SelectMany(line => line).Select(segment => segment.Text)).Replace("\u200B", string.Empty);
        PdfStandardFont font = ChooseNormal(options.DefaultFont);
        string fontResource = GetFontResourceName(font, null, font);
        ContentStreamBuilder anchor = WriteLogicalTextAnchor(sb, options, font, fontResource, x, y, logicalText, width, height);
        sb.Append("EMC\n");
        anchor.RestoreState();
    }

}
