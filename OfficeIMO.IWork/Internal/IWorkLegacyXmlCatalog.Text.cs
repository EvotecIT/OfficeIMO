using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkLegacyXmlCatalog {
    private readonly Dictionary<XElement, ulong> _listIdentifiers = new();
    private readonly Dictionary<XElement, (IWorkObjectIdentity Attachment, IWorkObjectIdentity Drawable)> _inlineObjects = new();

    internal void RegisterInlineObject(XElement reference, XElement attachment, XElement drawable) =>
        _inlineObjects.Add(reference, (Identity(attachment), Identity(drawable)));

    internal IWorkTextContent ReadText(XElement? body, IWorkParagraphStyle? defaultStyle = null) {
        var paragraphs = new List<IWorkTextParagraph>();
        int textOffset = 0;
        if (body != null) ReadParagraphs(body, paragraphs, defaultStyle ?? PlainParagraphStyle, ref textOffset);
        return new IWorkTextContent(paragraphs, isComplete: false, isTextComplete: false);
    }

    private void ReadParagraphs(XElement parent, List<IWorkTextParagraph> paragraphs,
        IWorkParagraphStyle defaultStyle, ref int textOffset) {
        foreach (XElement element in parent.Elements()) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            if (element.Name == Sf + "section" || element.Name == Sf + "layout") {
                ReadParagraphs(element, paragraphs, defaultStyle, ref textOffset);
            } else if (element.Name == Sf + "p") {
                _budget.AddTextItem();
                XElement? selected = Style(element, (string?)element.Attribute(Sf + "style"), "paragraphstyle");
                IWorkParagraphStyle style = selected == null ? defaultStyle : ParagraphStyle(selected);
                var runs = new List<IWorkTextRun>();
                ReadTextRuns(element, runs, style.TextStyle, null, ref textOffset);
                var list = ReadList(element, selected);
                bool pageBreak = element.Elements(Sf + "pgbr").Any();
                paragraphs.Add(new IWorkTextParagraph(runs, style, list.Identifier, list.Level,
                    list.Label, pageBreak ? IWorkParagraphBreakKind.Page : IWorkParagraphBreakKind.Paragraph,
                    listMarkerKind: list.Kind));
                textOffset++;
            } else if (element.Name != Sf + "page-start" && element.Name != Sf + "container-hint"
                && element.Name != Sf + "insertion-point") {
                AddLoss("IWORK_LEGACY_TEXT_STRUCTURE_UNSUPPORTED",
                    "An XML text body contains an unsupported block; that block is omitted.");
            }
        }
    }

    private void ReadTextRuns(XElement element, List<IWorkTextRun> runs, IWorkTextStyle style,
        string? hyperlink, ref int textOffset) {
        foreach (XNode node in element.Nodes()) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            if (node is XElement child) {
                if (child.Name == Sf + "span") {
                    XElement? selected = Style(child, (string?)child.Attribute(Sf + "style"), "characterstyle");
                    ReadTextRuns(child, runs, TextStyle(selected, style), hyperlink, ref textOffset);
                    continue;
                }
                if (child.Name == Sf + "link") {
                    string? target = (string?)child.Attribute("href") ?? (string?)child.Attribute(Sf + "href");
                    ReadTextRuns(child, runs, style, target, ref textOffset);
                    continue;
                }
                if (child.Name == Sf + "ghost-text") {
                    ReadTextRuns(child, runs, style, hyperlink, ref textOffset);
                    continue;
                }
                if (_inlineObjects.TryGetValue(child, out var inline)) {
                    _budget.AddTextItem();
                    _budget.AddTextCharacters(1);
                    runs.Add(new IWorkTextRun(string.Empty, style, hyperlink,
                        new IWorkInlineObject(textOffset++, inline.Attachment, inline.Drawable)));
                    continue;
                }
            }
            string? value = node switch {
                XText text => text.Value,
                XElement control when control.Name.Namespace == Sf => control.Name.LocalName switch {
                    "tab" => "\t",
                    "lnbr" or "crbr" or "intratopicbr" => "\n",
                    _ => null
                },
                _ => null
            };
            if (value == null) {
                if (node is XElement omitted && omitted.Name != Sf + "br" && omitted.Name != Sf + "pgbr"
                    && omitted.Name != Sf + "insertion-point")
                    AddLoss("IWORK_LEGACY_TEXT_INLINE_UNSUPPORTED",
                        "An XML inline field, attachment, or unsupported text element is omitted from its text storage.");
                continue;
            }
            _budget.AddTextItem();
            _budget.AddTextCharacters(value.Length);
            _budget.AddTextCharacters(style.FontName?.Length ?? 0);
            _budget.AddTextCharacters(hyperlink?.Length ?? 0);
            runs.Add(new IWorkTextRun(value, style, hyperlink));
            textOffset += value.Length;
        }
    }

    private (ulong? Identifier, int Level, string? Label, IWorkListMarkerKind Kind) ReadList(
        XElement paragraph, XElement? paragraphStyle) {
        var none = ((ulong?)null, -1, (string?)null, IWorkListMarkerKind.None);
        XElement? property = Property(paragraphStyle, "listStyle");
        XElement? style = Resolve(property?.Elements().FirstOrDefault(e => e.Name != Sf + "null"), "liststyle");
        if (style == null) return none;
        XElement[] types = Property(style, "listLabelTypes")?.Elements().FirstOrDefault()
            ?.Elements(Sf + "list-label-typeinfo").Take(10).ToArray() ?? Array.Empty<XElement>();
        string? levelText = (string?)paragraph.Attribute(Sf + "list-level");
        int level = 0;
        if (levelText != null && !int.TryParse(levelText, NumberStyles.None, CultureInfo.InvariantCulture, out level))
            return UnsupportedList();
        if (level < 0 || level >= types.Length || types.Length > 9) return UnsupportedList();
        XElement type = types[level];
        string? kind = (string?)type.Attribute(Sf + "type");
        if (kind == "none") return none;
        XElement? label = type.Element(Sf + "text-label");
        string? marker = (string?)label?.Attribute(Sf + "format");
        if (kind != "text" || (string?)label?.Attribute(Sf + "type") != "bullet"
            || string.IsNullOrEmpty(marker)) return UnsupportedList();
        if (!_listIdentifiers.TryGetValue(style, out ulong identifier)) {
            identifier = (ulong)_listIdentifiers.Count + 1;
            _listIdentifiers.Add(style, identifier);
        }
        _budget.AddTextCharacters(marker!.Length);
        bool unlabeledFirst = (string?)types[0].Attribute(Sf + "type") == "none";
        return (identifier, unlabeledFirst ? level - 1 : level, marker, IWorkListMarkerKind.Text);

        (ulong?, int, string?, IWorkListMarkerKind) UnsupportedList() {
            AddLoss("IWORK_LEGACY_LIST_UNSUPPORTED",
                "A selected XML list marker or level is outside the qualified text-marker profile.");
            return none;
        }
    }
}
