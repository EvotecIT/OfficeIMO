using System.Linq;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Table;

internal static class OfficeTableCellTextRunStyle
{
    internal static OfficeTextRunSpec[] Apply(OfficeTextRunSpec[] runs, OfficeTableCellStyle? style)
    {
        if (style?.HasTextStyle != true)
        {
            return runs;
        }

        return runs.Select(run => OfficeTextRunParser.NormalizeDerivedFields(new OfficeTextRunSpec
        {
            Text = run.Text,
            Kind = run.Kind,
            Bold = run.Bold || style.Bold,
            Italic = run.Italic || style.Italic,
            Underline = run.Underline || style.Underline || !string.IsNullOrWhiteSpace(style.UnderlineStyle),
            UnderlineStyle = run.UnderlineStyle ?? style.UnderlineStyle,
            Strike = run.Strike || style.Strike,
            Color = run.Color ?? style.TextColor,
            BackgroundColor = run.BackgroundColor,
            FontSize = run.FontSize ?? style.FontSize,
            FontName = run.FontName,
            Baseline = run.Baseline,
            LinkUri = run.LinkUri,
            LinkDestinationName = run.LinkDestinationName,
            LinkContents = run.LinkContents,
            TabLeader = run.TabLeader,
            TabAlignment = run.TabAlignment
        })).ToArray();
    }
}
