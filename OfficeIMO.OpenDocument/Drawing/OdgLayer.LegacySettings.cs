namespace OfficeIMO.OpenDocument;

// Older LibreOffice drawings store document-layer flags in their first saved view.
// Explicit ODF layer attributes take precedence; edits synchronize existing view masks.
internal static class OdgLegacyLayerSettings {
    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    internal static bool? Read(OdfDocument document, XElement layer, string name) {
        int index = LayerIndex(layer);
        if (index < 0) return null;
        XElement? item = Items(document, name).FirstOrDefault();
        if (item == null) return null;
        byte[] mask = Decode(item);
        return index / 8 < mask.Length && (mask[index / 8] & (1 << (index % 8))) != 0;
    }
    internal static void Write(OdfDocument document, XElement layer, string name, bool value) => WriteMasks(document, layer, (name, value));
    internal static void WriteDisplay(OdfDocument document, XElement layer, OdgLayerDisplay display) => WriteMasks(document, layer,
        ("VisibleLayers", display is OdgLayerDisplay.Always or OdgLayerDisplay.Screen),
        ("PrintableLayers", display is OdgLayerDisplay.Always or OdgLayerDisplay.Printer));
    private static void WriteMasks(OdfDocument document, XElement layer, params (string Name, bool Value)[] flags) {
        int index = LayerIndex(layer);
        if (index < 0) return;
        // Decode every affected mask before changing any saved view, including both display flags.
        var edits = flags.SelectMany(flag => Items(document, flag.Name).Select(item => {
            byte[] mask = Decode(item);
            if (index / 8 >= mask.Length) Array.Resize(ref mask, index / 8 + 1);
            if (flag.Value) mask[index / 8] |= (byte)(1 << (index % 8));
            else mask[index / 8] &= (byte)~(1 << (index % 8));
            return new { Item = item, Value = Convert.ToBase64String(mask) };
        })).ToList();
        foreach (var edit in edits) edit.Item.Value = edit.Value;
        if (edits.Count > 0) document.MarkPartDirty("settings.xml");
    }
    private static int LayerIndex(XElement layer) => layer.Parent?.Parent?.Name == OdfNamespaces.Office + "master-styles"
        ? layer.ElementsBeforeSelf(OdfNamespaces.Draw + "layer").Count() : -1;
    private static IEnumerable<XElement> Items(OdfDocument document, string name) {
        if (!document.Package.ContainsEntry("settings.xml")) return Enumerable.Empty<XElement>();
        return document.GetXml("settings.xml").Descendants(Config + "config-item-set")
            .Where(set => (string?)set.Attribute(Config + "name") == "ooo:view-settings")
            .Elements(Config + "config-item-map-indexed").Where(map => (string?)map.Attribute(Config + "name") == "Views")
            .Elements(Config + "config-item-map-entry").Elements(Config + "config-item")
            .Where(item => (string?)item.Attribute(Config + "name") == name && (string?)item.Attribute(Config + "type") == "base64Binary");
    }
    private static byte[] Decode(XElement item) {
        if (item.Value.Length > 131072) throw new InvalidDataException("Legacy layer state exceeds the supported mask size.");
        try { return Convert.FromBase64String(item.Value); }
        catch (FormatException exception) { throw new InvalidDataException("Legacy layer state is not valid base64.", exception); }
    }
}
