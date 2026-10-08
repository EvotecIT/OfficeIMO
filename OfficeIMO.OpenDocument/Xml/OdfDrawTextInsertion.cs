namespace OfficeIMO.OpenDocument;

/// <summary>Keeps Draw text before trailing enhanced geometry in native child order.</summary>
internal static class OdfDrawTextInsertion {
    internal static void Append(XElement container, XElement text) {
        XElement? geometry = container.Element(OdfNamespaces.Draw + "enhanced-geometry");
        if (geometry == null) container.Add(text);
        else geometry.AddBeforeSelf(text);
    }
}
