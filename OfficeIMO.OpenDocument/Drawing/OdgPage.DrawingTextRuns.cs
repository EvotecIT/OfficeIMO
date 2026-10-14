using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static List<OfficeRichTextRun> ReadDrawingRuns(OdfTextParagraph paragraph, OdgShape shape, ref int remaining, ref int visitedNodes, HashSet<string> losses, DrawingFieldResolver fields) {
        var runs = new List<OfficeRichTextRun>();
        int visited = visitedNodes;
        OdfTextCodec.TextSnapshot decoded;
        try {
            decoded = OdfTextCodec.Snapshot(paragraph.Element, remaining, () => {
                if (++visited > OdfTextTraversal.MaximumVisitedElements)
                    throw new NotSupportedException("Inline text traversal exceeds the projection node limit.");
            }, fields.Resolve);
        } finally { visitedNodes = visited; }
        remaining -= decoded.SourceCharacters;
        Read(paragraph.Element, paragraph, null, 0);
        return runs;

        void Read(XElement parent, OdfTextContent format, string? link, int depth) {
            if (depth > OdfTextTraversal.MaximumContainerDepth) throw new NotSupportedException("Inline text nesting exceeds the projection limit.");
            foreach (XNode node in parent.Nodes()) {
                if (node is XElement inline && (inline.Name == OdfNamespaces.Text + "span" || inline.Name == OdfNamespaces.Text + "a")) {
                    string? target = link;
                    if (inline.Name == OdfNamespaces.Text + "a") {
                        if (!OfficeDrawingLinkPolicy.TryNormalize((string?)inline.Attribute(OdfNamespaces.XLink + "href"), out string uri)) {
                            target = null; losses.Add("hyperlink-target");
                        } else target = uri;
                    }
                    Read(inline, new OdfTextRun(shape.Document, inline, shape.Element), target, depth + 1);
                    continue;
                }
                if (node is XElement element) {
                    if (OdfTextCodec.IsNonVisibleTextElement(element)) continue;
                    if (OdfTextField.IsField(element.Name)) { if (element.HasElements) losses.Add("field-content"); }
                    else if (element.Name != OdfNamespaces.Text + "s" && element.Name != OdfNamespaces.Text + "tab" && element.Name != OdfNamespaces.Text + "line-break")
                        losses.Add("unmapped-inline");
                }
                string text = decoded.ReadNode(node);
                if (text.Length == 0) continue;
                if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) throw new NotSupportedException("Text exceeds the shared styled-run layout limit.");
                OfficeRichTextRun run = CreateDrawingRun(text, format, losses); run.LinkUri = link; runs.Add(run);
            }
        }

    }

    private static OfficeRichTextRun CreateDrawingRun(string text, OdfTextContent source, HashSet<string> losses) {
        double size = OdfTextStyleResolver.ResolveFontSize(source.Styles, losses);
        OfficeTextDecorationStyle underline = source.UnderlineStyle switch {
            null => source.Underline == true ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            OdfTextDecorationStyle.None => OfficeTextDecorationStyle.None,
            OdfTextDecorationStyle.Solid => source.UnderlineType == OdfTextDecorationType.Double ? OfficeTextDecorationStyle.Double : OfficeTextDecorationStyle.Single,
            OdfTextDecorationStyle.Dotted => OfficeTextDecorationStyle.Dotted,
            OdfTextDecorationStyle.Dash => OfficeTextDecorationStyle.Dashed,
            OdfTextDecorationStyle.Wave => OfficeTextDecorationStyle.Wavy,
            _ => OfficeTextDecorationStyle.Single
        };
        if (source.Underline == false) underline = OfficeTextDecorationStyle.None;
        if (source.UnderlineType == OdfTextDecorationType.Double && source.UnderlineStyle is not (null or OdfTextDecorationStyle.Solid or OdfTextDecorationStyle.None)) losses.Add("underline-pattern");
        OdfTextDecorationStyle? strike = source.Resolve(style => style.LineThroughStyle);
        OdfTextDecorationType? strikeType = source.Resolve(style => style.LineThroughType);
        OfficeTextDecorationStyle strikeStyle = strike switch {
            OdfTextDecorationStyle.Dotted => OfficeTextDecorationStyle.Dotted,
            OdfTextDecorationStyle.Dash => OfficeTextDecorationStyle.Dashed,
            OdfTextDecorationStyle.Wave => OfficeTextDecorationStyle.Wavy,
            _ => strikeType == OdfTextDecorationType.Double ? OfficeTextDecorationStyle.Double : OfficeTextDecorationStyle.Single
        };
        if (strike is OdfTextDecorationStyle.LongDash or OdfTextDecorationStyle.DotDash or OdfTextDecorationStyle.DotDotDash) losses.Add("line-through-pattern");
        if (source.UnderlineStyle is OdfTextDecorationStyle.LongDash or OdfTextDecorationStyle.DotDash or OdfTextDecorationStyle.DotDotDash) losses.Add("underline-pattern");
        if (source.SmallCaps == true) losses.Add("small-caps");
        if (source.TextPosition is OdfTextPosition.Superscript or OdfTextPosition.Subscript) losses.Add("baseline-metrics");
        OfficeTextCase textCase = source.TextTransform switch {
            OdfTextTransform.Uppercase => OfficeTextCase.Uppercase,
            OdfTextTransform.Lowercase => OfficeTextCase.Lowercase,
            _ => OfficeTextCase.None
        };
        if (source.TextTransform == OdfTextTransform.Capitalize) losses.Add("capitalization");
        if (source.TextTransform is OdfTextTransform.Uppercase or OdfTextTransform.Lowercase) losses.Add("case-language-context");
        var seen = new HashSet<XName>();
        foreach (OdfStyle style in source.Styles) {
            foreach (XAttribute attribute in style.TextProperties?.Attributes() ?? Enumerable.Empty<XAttribute>()) {
                if (!seen.Add(attribute.Name)) continue;
                if (attribute.Name == OdfNamespaces.Style + "text-outline" && attribute.Value == "true") losses.Add("text-outline");
                if (attribute.Name == OdfNamespaces.Fo + "text-shadow" && attribute.Value != "none") losses.Add("text-shadow");
                if (attribute.Name == OdfNamespaces.Fo + "letter-spacing" && attribute.Value is not ("normal" or "0pt" or "0cm")) losses.Add("letter-spacing");
                if (attribute.Name == OdfNamespaces.Style + "letter-kerning" && attribute.Value == "false") losses.Add("kerning-policy");
                if (attribute.Name == OdfNamespaces.Style + "font-relief" && attribute.Value != "none") losses.Add("font-relief");
                if (attribute.Name == OdfNamespaces.Style + "text-scale" && attribute.Value != "100%") losses.Add("text-scale");
                if (attribute.Name == OdfNamespaces.Style + "text-line-through-text" && attribute.Value.Length > 0) losses.Add("line-through-glyph");
                if (underline != OfficeTextDecorationStyle.None) ReportDecoration("underline");
                if (source.StrikeThrough == true) ReportDecoration("line-through");
                if ((attribute.Name.LocalName.EndsWith("-complex", StringComparison.Ordinal) && text.Any(c => c >= '\u0590' && c <= '\u0DFF')) ||
                    (attribute.Name.LocalName.EndsWith("-asian", StringComparison.Ordinal) && text.Any(c => c >= '\u2E80' && c <= '\uA4FF'))) losses.Add("script-specific-formatting");

                void ReportDecoration(string decoration) {
                    if (attribute.Name == OdfNamespaces.Style + ("text-" + decoration + "-mode") && attribute.Value != "continuous")
                        losses.Add(decoration + "-mode");
                    if (attribute.Name == OdfNamespaces.Style + ("text-" + decoration + "-width") && attribute.Value != "auto")
                        losses.Add(decoration + "-width");
                    if (attribute.Name == OdfNamespaces.Style + ("text-" + decoration + "-color") && attribute.Value != "font-color" &&
                        !string.Equals(attribute.Value, source.Color?.ToString() ?? "#000000", StringComparison.OrdinalIgnoreCase))
                        losses.Add(decoration + "-color");
                }
            }
        }
        var result = new OfficeRichTextRun(text, size, DrawingTextColor(source, losses),
            source.Bold ?? false, source.Italic ?? false, underline != OfficeTextDecorationStyle.None,
            source.FontFamily ?? "Arial", source.StrikeThrough ?? false, ToColor(source.BackgroundColor), underlineStyle: underline,
            strikethroughStyle: source.StrikeThrough == true ? strikeStyle : OfficeTextDecorationStyle.None,
            baseline: source.TextPosition switch { OdfTextPosition.Superscript => OfficeTextBaseline.Superscript,
                OdfTextPosition.Subscript => OfficeTextBaseline.Subscript, _ => OfficeTextBaseline.Normal });
        return result.WithTextCase(textCase, CultureInfo.InvariantCulture);
    }

}
