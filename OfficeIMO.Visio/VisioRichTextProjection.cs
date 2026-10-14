using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Projects unchanged native character runs onto the shared drawing text model.</summary>
internal sealed partial class VisioRichTextProjection {
    internal IReadOnlyList<OfficeRichTextRun> Runs { get; }
    internal OfficeTextAlignment Alignment { get; }
    internal IReadOnlyList<OfficeRichTextParagraph> Paragraphs { get; }
    internal OfficeTextAlignment RenderAlignment => Paragraphs.Count > 0 ? OfficeTextAlignment.Left : Alignment;

    private VisioRichTextProjection(IReadOnlyList<OfficeRichTextRun> runs, OfficeTextAlignment alignment,
        IReadOnlyList<OfficeRichTextParagraph>? paragraphs = null) {
        Runs = runs;
        Alignment = alignment;
        Paragraphs = paragraphs ?? Array.Empty<OfficeRichTextParagraph>();
    }

    internal static VisioRichTextProjection? Create(VisioPage page, VisioShape shape, double scale,
        CancellationToken cancellationToken, VisioNativeTextStyleResolver? resolver = null) {
        if (shape.Text?.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters) return null;
        bool explicitStyle = !string.IsNullOrEmpty(shape.NativeStyleReferences?.TextStyle);
        bool deleted = HasDeletedTextSection(shape.PreservedNonGeometrySections, shape.CharacterSectionSource, shape.ParagraphSectionSource);
        bool inherited = explicitStyle || shape.MasterShape?.NativeStyleReferences != null || deleted;
        resolver ??= inherited ? new VisioNativeTextStyleResolver(page.OwnerDocument, cancellationToken) : null;
        if (inherited && !deleted && resolver!.UsesUnformattedBase(shape.NativeStyleReferences?.TextStyle))
            inherited = false;
        if (inherited) {
            resolver ??= new VisioNativeTextStyleResolver(page.OwnerDocument, cancellationToken);
            bool replacement = !string.Equals(shape.Text, shape.PreservedTextValue, StringComparison.Ordinal);
            return Create(page.OwnerDocument, shape.Text, shape.Text,
                replacement || shape.PreservedTextElement == null ? PlainText(shape.Text) : shape.PreservedTextElement, null,
                resolver.Resolve(shape, true, replacement), resolver.Resolve(shape, false, replacement), scale, 10D, cancellationToken, true);
        }
        return Create(page.OwnerDocument, shape.Text,
            shape.PreservedTextValue, shape.PreservedTextElement, shape.TextStyle,
            TextSection(shape, character: true) ??
                (explicitStyle || shape.MasterShape == null ? null : TextSection(shape.MasterShape, character: true)),
            TextSection(shape, character: false) ??
                (explicitStyle || shape.MasterShape == null ? null : TextSection(shape.MasterShape, character: false)),
            scale, 10D, cancellationToken);
    }

    internal static VisioRichTextProjection? Create(VisioPage page, VisioConnector connector, double scale,
        CancellationToken cancellationToken, VisioNativeTextStyleResolver? resolver = null) {
        if (connector.Label?.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters) return null;
        bool deleted = HasDeletedTextSection(connector.PreservedNonGeometrySections, connector.CharacterSectionSource, connector.ParagraphSectionSource);
        bool inherited = !string.IsNullOrEmpty(connector.NativeStyleReferences?.TextStyle) || deleted;
        resolver ??= inherited ? new VisioNativeTextStyleResolver(page.OwnerDocument, cancellationToken) : null;
        if (inherited && !deleted && resolver!.UsesUnformattedBase(connector.NativeStyleReferences?.TextStyle)) inherited = false;
        if (inherited) {
            resolver ??= new VisioNativeTextStyleResolver(page.OwnerDocument, cancellationToken);
            bool replacement = !string.Equals(connector.Label, connector.PreservedTextValue, StringComparison.Ordinal);
            return Create(page.OwnerDocument, connector.Label, connector.Label,
                replacement || connector.PreservedTextElement == null ? PlainText(connector.Label) : connector.PreservedTextElement, null,
                resolver.Resolve(connector, true, replacement), resolver.Resolve(connector, false, replacement), scale, 9D, cancellationToken, true);
        }
        return Create(page.OwnerDocument, connector.Label,
            connector.PreservedTextValue, connector.PreservedTextElement, connector.TextStyle,
            TextSection(connector.PreservedNonGeometrySections,
                connector.CharacterSectionSource, connector.TextStyle, character: true),
            TextSection(connector.PreservedNonGeometrySections,
                connector.ParagraphSectionSource, connector.TextStyle, character: false),
            scale, 9D, cancellationToken);
    }

    private static XElement PlainText(string? text) => new(XName.Get("Text", VisioDocument.VisioNamespace), text ?? string.Empty);

    // Captured single rows and native-only multirow sections share the same edit
    // view. Rebind detached master font identities on a clone, never the source.
    private static XElement? TextSection(VisioShape shape, bool character) =>
        VisioDocument.GetMasterFontRenderSection(shape,
            TextSection(shape.PreservedNonGeometrySections,
                character ? shape.CharacterSectionSource : shape.ParagraphSectionSource,
                shape.TextStyle, character));

    private static XElement? TextSection(IEnumerable<XElement> sections, VisioTextSectionSource? source,
        VisioTextStyle? style, bool character) {
        XElement? native = character ? Character(sections) : Paragraph(sections);
        if (native == null) {
            // A plain typed label needs no run projection. Single captured
            // paragraphs containing only alignment still use the plain renderer;
            // extra native cells carry bullets, indentation or spacing.
            if (source == null || (!character && !source.Source.Descendants()
                    .Any(cell => cell.Name.LocalName == "Cell" && (string?)cell.Attribute("N") != "HorzAlign"))) return null;
        }
        return VisioDocument.GetRenderTextSection(sections, source, style, character);
    }

    private static XElement? Character(IEnumerable<XElement>? sections) => sections?.FirstOrDefault(section =>
        (string?)section.Attribute("N") is "Character" or "Char");

    private static XElement? Paragraph(IEnumerable<XElement>? sections) => sections?.FirstOrDefault(section =>
        (string?)section.Attribute("N") is "Paragraph" or "Para");

    private static bool HasDeletedTextSection(IEnumerable<XElement> sections, VisioTextSectionSource? character, VisioTextSectionSource? paragraph) =>
        VisioNativeTextStyleResolver.HasDeletion(Character(sections) ?? character?.Source) ||
        VisioNativeTextStyleResolver.HasDeletion(Paragraph(sections) ?? paragraph?.Source);

    private static VisioRichTextProjection? Create(VisioDocument? document, string? text,
        string? preservedValue, XElement? element, VisioTextStyle? style, XElement? characters,
        XElement? paragraph, double scale, double defaultSize, CancellationToken cancellationToken, bool effectiveStyle = false) {
        // Replacing the plain label deliberately replaces its native run boundaries.
        if (element == null || !string.Equals(text, preservedValue, StringComparison.Ordinal)) return null;
        if (text != null && text.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters) return null;
        XNamespace ns = element.Name.Namespace;
        var rows = new Dictionary<string, XElement>(StringComparer.Ordinal);
        int position = 0;
        foreach (XElement row in characters?.Elements(ns + "Row") ?? Enumerable.Empty<XElement>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (rows.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) return null;
            rows[VisioNativeTextStyleResolver.RowIndex((string?)row.Attribute("IX"), position)] = row;
            position++;
        }
        OfficeTextAlignment alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment(style?.HorizontalAlignment);
        if (paragraph != null) {
            VisioRichTextProjection? projection = CreateParagraphProjection(document, element, rows, paragraph, style,
                scale, defaultSize, alignment, cancellationToken);
            if (projection != null) return projection;
        }
        if (rows.Count < (effectiveStyle ? 1 : 2)) return null;
        string index = "0";
        var runs = new List<OfficeRichTextRun>();
        foreach (XNode node in element.DescendantNodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node is XElement marker && marker.Name == ns + "cp") {
                index = VisioNativeTextStyleResolver.RowIndex((string?)marker.Attribute("IX"));
            } else if (node is XText value && value.Value.Length > 0) {
                if (!rows.TryGetValue(index, out XElement? row) || runs.Count >= 4096) return null;
                runs.Add(CreateRun(value.Value, row, document, style, scale, defaultSize));
            }
        }
        if (runs.Count == 0) return null;
        // Unsupported spacing or geometry does not discard a single row's supported alignment.
        List<XElement>? paragraphs = paragraph?.Elements(ns + "Row").Take(2).ToList();
        XElement? para = paragraphs?.Count == 1 ? paragraphs[0] : null;
        if (para != null && TryInt(Cell(para, "HorzAlign"), out int align) && align >= 0 && align <= 3)
            alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment((VisioTextHorizontalAlignment)align);
        return new VisioRichTextProjection(runs, alignment);
    }

    private static OfficeRichTextRun CreateRun(string text, XElement row, VisioDocument? document,
        VisioTextStyle? fallback, double scale, double defaultSize) {
        int bits = TryInt(Cell(row, "Style"), out int parsedBits) ? parsedBits :
            (fallback?.Bold == true ? 1 : 0) | (fallback?.Italic == true ? 2 : 0) | (fallback?.Underline == true ? 4 : 0);
        double points = fallback?.Size ?? defaultSize;
        if (double.TryParse(Cell(row, "Size"), NumberStyles.Float, CultureInfo.InvariantCulture, out double inches)
            && inches > 0D && !double.IsInfinity(inches)) points = inches * 72D;
        int casing = TryInt(Cell(row, "Case"), out int parsedCase) ? parsedCase : (int?)fallback?.Capitalization ?? 0;
        if (casing == 1) text = OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);
        else if (casing == 2) text = OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Capitalize, CultureInfo.InvariantCulture);
        OfficeTextBaseline baseline = fallback?.Baseline ?? OfficeTextBaseline.Normal;
        if (TryInt(Cell(row, "Pos"), out int position) && position >= 0 && position <= 2) baseline = (OfficeTextBaseline)position;
        bool underline = (bits & 4) != 0;
        OfficeTextDecorationStyle underlineStyle = ResolveDecoration(row, "DblUnderline", null,
            underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None);
        OfficeTextDecorationStyle strike = ResolveDecoration(row, "DoubleStrikethrough", "Strikethru",
            fallback?.StrikethroughStyle ?? OfficeTextDecorationStyle.None);
        OfficeColor color = Color(Cell(row, "Color"), document, fallback?.Color ?? OfficeColor.FromRgb(17, 24, 39));
        if (double.TryParse(Cell(row, "ColorTrans"), NumberStyles.Float, CultureInfo.InvariantCulture, out double transparency)
            && !double.IsNaN(transparency)) {
            double alpha = color.A * (1D - Math.Max(0D, Math.Min(1D, transparency)));
            color = OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(alpha));
        }
        return new OfficeRichTextRun(text, points * scale / 72D, color,
            (bits & 1) != 0, (bits & 2) != 0, underline,
            Font(Cell(row, "Font"), document, fallback?.FontFamily), strike != OfficeTextDecorationStyle.None,
            underlineStyle: underlineStyle, strikethroughStyle: strike, baseline: baseline);
    }

    private static OfficeTextDecorationStyle ResolveDecoration(XElement row, string doubleName, string? singleName,
        OfficeTextDecorationStyle fallback) {
        if (TryInt(Cell(row, doubleName), out int doubleValue) && doubleValue != 0) return OfficeTextDecorationStyle.Double;
        if (singleName != null && TryInt(Cell(row, singleName), out int singleValue))
            return singleValue == 0 ? OfficeTextDecorationStyle.None : OfficeTextDecorationStyle.Single;
        return fallback;
    }

    private static string? Cell(XElement row, string name) => (string?)row.Elements()
        .FirstOrDefault(cell => cell.Name.LocalName == "Cell" && (string?)cell.Attribute("N") == name)?.Attribute("V");

    private static bool TryInt(string? value, out int result) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);

    private static string Font(string? value, VisioDocument? document, string? fallback) {
        if (TryInt(value, out int id)) {
            string? name = document?.PreservedFaceNamesElements.FirstOrDefault(face =>
                TryInt((string?)face.Attribute("ID"), out int faceId) && faceId == id)?.Attribute("Name")?.Value;
            if (!string.IsNullOrWhiteSpace(name)) return name!.Length > 256 ? name.Substring(0, 256) : name;
        } else if (!string.IsNullOrWhiteSpace(value) && !string.Equals(value, "themed", StringComparison.OrdinalIgnoreCase))
            return value!.Length > 256 ? value.Substring(0, 256).Trim('"') : value.Trim('"');
        return string.IsNullOrWhiteSpace(fallback) ? "Aptos, Calibri, Arial, sans-serif" : fallback!;
    }

    private static OfficeColor Color(string? value, VisioDocument? document, OfficeColor fallback) {
        if (string.Equals(value, "themed", StringComparison.OrdinalIgnoreCase)) return fallback;
        if (TryInt(value, out int id)) {
            value = document?.PreservedColorsElements.FirstOrDefault(entry =>
                TryInt((string?)entry.Attribute("IX"), out int colorId) && colorId == id)?.Attribute("RGB")?.Value ?? value;
        }
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        try { return VisioHelpers.FromVisioColor(value!); }
        catch (FormatException) { return fallback; }
        catch (ArgumentException) { return fallback; }
        catch (OverflowException) { return fallback; }
    }
}
