using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private sealed class TabLeaderText : OdfTextContent {
        private readonly IReadOnlyList<OdfStyle> _styles;
        internal TabLeaderText(OdfTextParagraph paragraph, IReadOnlyList<OdfStyle> styles)
            : base(paragraph.Document, paragraph.Element, paragraph.Graphic, OdfStyleFamily.Text) => _styles = styles;
        internal override IEnumerable<OdfStyle> Styles => _styles;
    }

    private static OfficeTextTabStop ProjectStyledTabLeader(OfficeTextTabStop stop, OdfTextParagraph paragraph,
        OdfStyle tabStyle, string name, HashSet<string> losses) {
        OdfStyleRepository repository = paragraph.Document.Styles;
        OdfStyle? style = tabStyle.IsAutomatic ? repository.FindInPart(OdfStyleFamily.Text, name, tabStyle.PartPath) : repository.FindNamed(OdfStyleFamily.Text, name);
        if (style == null) throw new NotSupportedException("The referenced leader text style is missing in its original scope.");
        IReadOnlyList<OdfStyle> chain = repository.Resolve(style);
        if (!string.IsNullOrEmpty(chain[chain.Count - 1].ParentStyleName)) losses.Add("tab-leader-style-chain");
        var source = new TabLeaderText(paragraph, chain);
        OfficeRichTextRun mapped = CreateDrawingRun(stop.LeaderText!, source, losses);
        var size = OdfTextStyleResolver.ResolveFontSizeBasis(chain, losses);
        double? points = size.Points.HasValue ? size.Points.Value * size.Factor : null;
        bool hasBackground = chain.Any(s => s.TryGetTextBackgroundColor(out _));
        var formatting = new OfficeTextTabLeaderStyle(points, size.Points.HasValue ? 1 : size.Factor,
            ToColor(source.Color), source.FontFamily, source.Bold, source.Italic,
            source.Underline.HasValue || source.UnderlineStyle.HasValue || source.UnderlineType.HasValue ? mapped.UnderlineStyle : null,
            source.StrikeThrough.HasValue ? mapped.StrikethroughStyle : null,
            source.TextPosition.HasValue ? mapped.Baseline : null,
            hasBackground ? ToColor(source.BackgroundColor) ?? OfficeColor.Transparent : null, inheritOpacity: true, opacity: source.TextOpacity);
        // Display-time casing is compiled through the existing run owner. A mapping
        // that expands the native scalar beyond this glyph profile fails explicitly.
        return stop.WithLeader(mapped.Text).WithLeaderStyle(formatting);
    }
}
