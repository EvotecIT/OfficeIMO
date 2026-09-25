using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Materializes structurally supported ARIA tables for native document adapters.</summary>
internal static class HtmlRoleTableNormalizer {
    internal static void Normalize(
        IHtmlDocument document,
        Action<IElement, IElement>? registerNative = null,
        Action<IElement>? materializeCss = null,
        ISet<IElement>? materializedElements = null,
        Action<IElement>? onUnsupported = null) {
        // A conversion gets its own DOM clone. Work from the innermost table out so a
        // nested table is already native when its containing cell is moved below.
        IElement[] tables = document.QuerySelectorAll("[role]")
            .Where(element => element.LocalName != "table" && HtmlAccessibilitySemantics.HasRole(element, "table"))
            .ToArray();
        if (materializeCss != null) {
            var styledElements = new HashSet<IElement>();
            foreach (IElement table in tables.Where(CanNormalizeRoleTable)) {
                styledElements.Add(table);
                foreach (IElement descendant in table.QuerySelectorAll("*")) styledElements.Add(descendant);
            }
            foreach (IElement element in document.QuerySelectorAll("*")) {
                if (styledElements.Contains(element)) {
                    materializeCss(element);
                    materializedElements?.Add(element);
                }
            }
        }
        for (int index = tables.Length - 1; index >= 0; index--) {
            IElement table = tables[index];
            if (table.Parent == null) continue;
            if (!CanNormalizeRoleTable(table)) {
                onUnsupported?.Invoke(table);
                continue;
            }

            IElement nativeTable = document.CreateElement("table");
            CopyRoleTableAttributes(table, nativeTable);
            registerNative?.Invoke(table, nativeTable);
            IElement? implicitBody = null;
            foreach (IElement child in table.Children) {
                if (HtmlAccessibilitySemantics.HasRole(child, "rowgroup")) {
                    implicitBody = null;
                    IElement body = document.CreateElement("tbody");
                    CopyRoleTableAttributes(child, body);
                    registerNative?.Invoke(child, body);
                    foreach (IElement row in child.Children) {
                        body.AppendChild(CreateNativeRoleRow(document, row, registerNative));
                    }
                    nativeTable.AppendChild(body);
                } else {
                    if (implicitBody == null) {
                        implicitBody = document.CreateElement("tbody");
                        registerNative?.Invoke(table, implicitBody);
                        nativeTable.AppendChild(implicitBody);
                    }
                    implicitBody.AppendChild(CreateNativeRoleRow(document, child, registerNative));
                }
            }
            table.Parent.ReplaceChild(nativeTable, table);
        }
    }

    private static bool CanNormalizeRoleTable(IElement table) {
        if (!IsNeutralRoleContainer(table)) return false;
        bool hasRow = false;
        foreach (INode child in table.ChildNodes) {
            if (IsIgnorableRoleTableNode(child)) continue;
            if (child is not IElement element) return false;
            if (HtmlAccessibilitySemantics.HasRole(element, "rowgroup")) {
                if (!IsNeutralRoleContainer(element)) return false;
                foreach (INode row in element.ChildNodes) {
                    if (IsIgnorableRoleTableNode(row)) continue;
                    if (row is not IElement rowElement || !CanNormalizeRoleRow(rowElement)) return false;
                    hasRow = true;
                }
            } else if (CanNormalizeRoleRow(element)) {
                hasRow = true;
            } else {
                return false;
            }
        }
        return hasRow;
    }

    private static bool CanNormalizeRoleRow(IElement row) {
        if (!IsNeutralRoleContainer(row) || !HtmlAccessibilitySemantics.HasRole(row, "row")) return false;
        bool hasCell = false;
        foreach (INode child in row.ChildNodes) {
            if (IsIgnorableRoleTableNode(child)) continue;
            if (child is not IElement cell ||
                !(HtmlAccessibilitySemantics.HasRole(cell, "cell") ||
                  HtmlAccessibilitySemantics.HasRole(cell, "columnheader") ||
                  HtmlAccessibilitySemantics.HasRole(cell, "rowheader")) ||
                IsNativeTableStructure(cell)) return false;
            hasCell = true;
        }
        return hasCell;
    }

    private static bool IsIgnorableRoleTableNode(INode node) =>
        node.NodeType == NodeType.Comment ||
        (node.NodeType == NodeType.Text && string.IsNullOrWhiteSpace(node.TextContent));

    private static bool IsNeutralRoleContainer(IElement element) =>
        element.LocalName is "div" or "span";

    private static bool IsNativeTableStructure(IElement element) =>
        element.LocalName is "table" or "thead" or "tbody" or "tfoot" or "tr" or "td" or "th";

    private static IElement CreateNativeRoleRow(
        IHtmlDocument document, IElement row, Action<IElement, IElement>? registerNative) {
        IElement nativeRow = document.CreateElement("tr");
        CopyRoleTableAttributes(row, nativeRow);
        registerNative?.Invoke(row, nativeRow);
        foreach (IElement cell in row.Children.ToArray()) {
            bool isHeader = HtmlAccessibilitySemantics.HasRole(cell, "columnheader") ||
                            HtmlAccessibilitySemantics.HasRole(cell, "rowheader");
            IElement nativeCell = document.CreateElement(isHeader ? "th" : "td");
            CopyRoleTableAttributes(cell, nativeCell);
            registerNative?.Invoke(cell, nativeCell);
            nativeCell.RemoveAttribute("id");
            nativeCell.RemoveAttribute("name");
            nativeCell.RemoveAttribute("role");
            foreach (string span in new[] { "rowspan", "colspan" }) {
                if (!nativeCell.HasAttribute(span) && cell.HasAttribute("aria-" + span)) {
                    nativeCell.SetAttribute(span, cell.GetAttribute("aria-" + span));
                }
            }
            nativeCell.AppendChild(cell);
            nativeRow.AppendChild(nativeCell);
        }
        return nativeRow;
    }

    private static void CopyRoleTableAttributes(IElement source, IElement destination) {
        foreach (IAttr attribute in source.Attributes) {
            destination.SetAttribute(attribute.Name, attribute.Value);
        }
    }
}
