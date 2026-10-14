using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Maps typed value writes to the canonical writer's relative own-cell identities.</summary>
internal static class VisioNativeCellAssignments {
    internal static ISet<string> For(VisioShape shape, VisioDocument? document = null) {
        var addresses = ForData(shape.ShapeData, shape.ShapeDataSectionName);
        foreach (VisioUserCell row in shape.UserCells) {
            if (row.ValueAssigned) addresses.Add(Address("User", row.Name, row.RowIndex, "Value"));
            if (row.PromptAssigned) addresses.Add(Address("User", row.Name, row.RowIndex, "Prompt"));
        }
        AddTextBackgroundAssignments(addresses, shape.TextStyle);
        AddFontAssignments(addresses, shape.TextStyle, shape.CharacterSectionSource?.Source);
        AddHyperlinkAssignments(addresses, shape.Hyperlinks, VisioHyperlinkRowNames.Inherited(shape, document));
        if (shape.NativeLayerMembership?.ProducerStateReplaced == true) addresses.Add("Cell[N=LayerMember]");
        return addresses;
    }

    internal static ISet<string> For(VisioConnector connector, VisioDocument? document = null) {
        var addresses = ForData(connector.ShapeData, connector.ShapeDataSectionName);
        AddTextBackgroundAssignments(addresses, connector.TextStyle);
        AddFontAssignments(addresses, connector.TextStyle, connector.CharacterSectionSource?.Source);
        AddHyperlinkAssignments(addresses, connector.Hyperlinks, VisioHyperlinkRowNames.Inherited(connector, document));
        if (connector.NativeLayerMembership?.ProducerStateReplaced == true) addresses.Add("Cell[N=LayerMember]");
        return addresses;
    }

    private static void AddTextBackgroundAssignments(ISet<string> addresses, VisioTextStyle? style) {
        if (style?.BackgroundColorAssigned == true) addresses.Add("Cell[N=TextBkgnd]");
        if (style?.BackgroundTransparencyAssigned == true) addresses.Add("Cell[N=TextBkgndTrans]");
    }

    internal static bool IsTextBackground(string address) => address == "Cell[N=TextBkgnd]" || address == "Cell[N=TextBkgndTrans]";

    private static void AddFontAssignments(ISet<string> addresses, VisioTextStyle? style, XElement? section) {
        if (style?.FontFamilyAssigned != true || section == null) return;
        foreach (XElement cell in VisioFontWireCodec.FontCells(section).Where(cell => (string?)cell.Attribute("N") == "Font"))
            addresses.Add(string.Join("/", cell.AncestorsAndSelf().TakeWhile(element => !ReferenceEquals(element, section.Parent))
                .Reverse().Select(VisioNativeCellMetadata.Step)));
    }

    internal static bool ReplacesProducerState(string address) => IsTextBackground(address) || address == "Cell[N=LayerMember]" ||
        address.EndsWith("/Cell[N=Font]", System.StringComparison.Ordinal) ||
        address.StartsWith("Section[N=Layer]/", System.StringComparison.Ordinal) ||
        address.StartsWith("Section[N=Hyperlink]/", System.StringComparison.Ordinal);

    private static void AddHyperlinkAssignments(ISet<string> addresses, IList<VisioHyperlink> hyperlinks, IList<VisioHyperlink>? inherited = null) {
        string?[] names = VisioHyperlinkRowNames.Create(hyperlinks, inherited);
        for (int i = 0; i < hyperlinks.Count; i++) {
            VisioHyperlink hyperlink = hyperlinks[i];
            foreach (string cellName in hyperlink.AssignedValues)
                addresses.Add(Address("Hyperlink", names[i], hyperlink.RowIndex, cellName));
        }
    }

    internal static ISet<string> ForLayers(IReadOnlyList<VisioLayer> layers) {
        var addresses = new HashSet<string>(System.StringComparer.Ordinal);
        int[] indexes = VisioLayerRowIndexes.Create(layers);
        for (int i = 0; i < layers.Count; i++)
            foreach (string cellName in layers[i].AssignedValues)
                addresses.Add(Address("Layer", null, indexes[i], cellName));
        return addresses;
    }

    private static HashSet<string> ForData(IEnumerable<VisioShapeDataRow> rows, string sectionName) {
        var addresses = new HashSet<string>(System.StringComparer.Ordinal);
        foreach (VisioShapeDataRow row in rows)
            foreach (string cellName in row.AssignedValues)
                addresses.Add(Address(sectionName, row.Name, row.RowIndex, cellName));
        return addresses;
    }

    private static string Address(string sectionName, string? rowName, int? rowIndex, string cellName) {
        XNamespace ns = VisioShapeSheetSection.VisioNamespace;
        var row = new XElement(ns + "Row");
        row.SetAttributeValue("N", rowName);
        row.SetAttributeValue("IX", rowIndex);
        var path = new[] {
            new XElement(ns + "Section", new XAttribute("N", sectionName)), row,
            new XElement(ns + "Cell", new XAttribute("N", cellName))
        };
        return string.Join("/", path.Select(VisioNativeCellMetadata.Step));
    }
}
