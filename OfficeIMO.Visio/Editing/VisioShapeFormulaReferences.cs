using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Remaps local numeric Sheet references in copied diagram graphs.</summary>
internal static class VisioShapeFormulaReferences {
    internal static void RemapPageSheet(VisioPage page, IReadOnlyDictionary<string, string> ids) {
        VisioDocument.ValidateCopiedNativePageSheetMetadata(page);
        RemapElements(page.PreservedPageSheetCells, ids);
        if (page.PageSheetLengthCells != null) RemapElements(new[] { page.PageSheetLengthCells.Source }, ids);
        RemapElements(page.PreservedPageSheetSections, ids);
        foreach (VisioLayer layer in page.Layers) {
            RemapElements(layer.PreservedKnownCells.Values, ids);
            RemapElements(layer.PreservedCells, ids);
        }
        page.NativePageSheetMetadata?.RemapFormulas(ids);
    }

    internal static void Remap(VisioShape shape, IReadOnlyDictionary<string, string> ids) {
        string? Formula(string? value) => Rewrite(value, ids);
        void Elements(IEnumerable<XElement> elements) => RemapElements(elements, ids);
        shape.RelationshipsFormula = Formula(shape.RelationshipsFormula);
        Elements(shape.PreservedCellElements);
        if (shape.NativeLayerMembership != null) Elements(new[] { shape.NativeLayerMembership.Cell });
        Elements(shape.PreservedGeometrySections);
        Elements(shape.PreservedNonGeometrySections);
        Elements(shape.PreservedDataRows);
        Elements(shape.PreservedShapeChildren.Where(entry => entry.RawElement != null).Select(entry => entry.RawElement!));
        if (shape.TextStyle?.NativeBackgroundColorCell is XElement shapeBackground) Elements(new[] { shapeBackground });
        if (shape.TextStyle?.NativeBackgroundTransparencyCell is XElement shapeTransparency) Elements(new[] { shapeTransparency });
        if (shape.PreservedTextElement != null) Elements(new[] { shape.PreservedTextElement });
        if (shape.CharacterSectionSource != null) Elements(new[] { shape.CharacterSectionSource.Source });
        if (shape.ParagraphSectionSource != null) Elements(new[] { shape.ParagraphSectionSource.Source });
        foreach (VisioUserCell row in shape.UserCells) {
            row.RemapFormulas(ids);
            Elements(row.PreservedCells);
        }
        RemapHyperlinks(shape.Hyperlinks, ids);
        RemapShapeData(shape.ShapeData, ids);
        shape.NativeCellMetadata?.RemapFormulas(ids);
    }

    internal static void Remap(VisioConnector connector, IReadOnlyDictionary<string, string> ids) {
        void Elements(IEnumerable<XElement> elements) => RemapElements(elements, ids);
        Elements(connector.PreservedCellElements);
        if (connector.NativeLayerMembership != null) Elements(new[] { connector.NativeLayerMembership.Cell });
        Elements(connector.PreservedEndpointCellElements.Values);
        Elements(connector.PreservedGeometrySections);
        Elements(connector.PreservedNonGeometrySections);
        Elements(connector.PreservedDataRows);
        Elements(connector.PreservedShapeChildren.Where(entry => entry.RawElement != null).Select(entry => entry.RawElement!));
        if (connector.TextStyle?.NativeBackgroundColorCell is XElement connectorBackground) Elements(new[] { connectorBackground });
        if (connector.TextStyle?.NativeBackgroundTransparencyCell is XElement connectorTransparency) Elements(new[] { connectorTransparency });
        if (connector.PreservedTextElement != null) Elements(new[] { connector.PreservedTextElement });
        if (connector.CharacterSectionSource != null) Elements(new[] { connector.CharacterSectionSource.Source });
        if (connector.ParagraphSectionSource != null) Elements(new[] { connector.ParagraphSectionSource.Source });
        RemapHyperlinks(connector.Hyperlinks, ids);
        RemapShapeData(connector.ShapeData, ids);
        connector.NativeCellMetadata?.RemapFormulas(ids);
    }

    private static void RemapElements(IEnumerable<XElement> elements, IReadOnlyDictionary<string, string> ids) {
        foreach (XAttribute formula in elements.SelectMany(element => element.DescendantsAndSelf()).Attributes("F"))
            formula.Value = Rewrite(formula.Value, ids)!;
    }

    private static void RemapHyperlinks(IEnumerable<VisioHyperlink> rows, IReadOnlyDictionary<string, string> ids) {
        foreach (VisioHyperlink row in rows) {
            RemapElements(row.PreservedCells, ids); RemapElements(row.PreservedKnownCells.Values, ids);
        }
    }

    private static void RemapShapeData(IEnumerable<VisioShapeDataRow> rows, IReadOnlyDictionary<string, string> ids) {
        string? Formula(string? value) => Rewrite(value, ids);
        foreach (VisioShapeDataRow row in rows) {
            row.ValueFormula = Formula(row.ValueFormula); row.LabelFormula = Formula(row.LabelFormula);
            row.PromptFormula = Formula(row.PromptFormula); row.TypeFormula = Formula(row.TypeFormula);
            row.FormatFormula = Formula(row.FormatFormula); row.SortKeyFormula = Formula(row.SortKeyFormula);
            row.InvisibleFormula = Formula(row.InvisibleFormula); row.VerifyFormula = Formula(row.VerifyFormula);
            row.DataLinkedFormula = Formula(row.DataLinkedFormula); row.CalendarFormula = Formula(row.CalendarFormula);
            row.LangIdFormula = Formula(row.LangIdFormula);
            RemapElements(row.PreservedCells, ids); RemapElements(row.PreservedKnownCells.Values, ids);
        }
    }

    internal static string NormalizeSheetId(string id) => uint.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint numeric)
        ? numeric.ToString(CultureInfo.InvariantCulture) : id;

    internal static string? Rewrite(string? formula, IReadOnlyDictionary<string, string> ids) {
        if (string.IsNullOrEmpty(formula)) return formula;
        var result = new StringBuilder(formula!.Length);
        bool quoted = false;
        for (int i = 0; i < formula.Length;) {
            char c = formula[i];
            if (c == '"') { quoted = !quoted; result.Append(c); i++; continue; }
            if (!quoted && (i == 0 || !char.IsLetterOrDigit(formula[i - 1]) && formula[i - 1] != '_' && formula[i - 1] != '!') &&
                i + 6 < formula.Length && string.Compare(formula, i, "Sheet.", 0, 6, StringComparison.OrdinalIgnoreCase) == 0) {
                int end = i + 6;
                while (end < formula.Length && formula[end] >= '0' && formula[end] <= '9') end++;
                if (end > i + 6 && end < formula.Length && formula[end] == '!' &&
                    ids.TryGetValue(NormalizeSheetId(formula.Substring(i + 6, end - i - 6)), out string? replacement)) {
                    result.Append(formula, i, 6).Append(replacement).Append('!'); i = end + 1; continue;
                }
            }
            result.Append(c); i++;
        }
        return result.ToString();
    }
}
