namespace OfficeIMO.OpenDocument;

internal static class OdfXmlCodec {
    internal static XDocument Load(byte[] bytes, string partPath, long maxCharacters, int maxDepth) {
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            XmlReaderSettings settings = CreateReaderSettings(maxCharacters);
            using (XmlReader depthReader = XmlReader.Create(stream, settings)) {
                while (depthReader.Read()) {
                    if (depthReader.Depth > maxDepth) throw new InvalidDataException($"OpenDocument XML part '{partPath}' exceeds MaxXmlDepth ({maxDepth}).");
                }
            }
            stream.Position = 0;
            settings = CreateReaderSettings(maxCharacters);
            using XmlReader reader = XmlReader.Create(stream, settings);
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        } catch (InvalidDataException) {
            throw;
        } catch (Exception ex) when (ex is XmlException || ex is InvalidOperationException) {
            throw new InvalidDataException($"OpenDocument XML part '{partPath}' is invalid.", ex);
        }
    }

    private static XmlReaderSettings CreateReaderSettings(long maxCharacters) => new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxCharacters,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                IgnoreWhitespace = false,
                CloseInput = false
            };

    internal static byte[] Save(XDocument document) {
        using var stream = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false,
            NewLineHandling = NewLineHandling.None,
            CloseOutput = false
        };
        using (XmlWriter writer = XmlWriter.Create(stream, settings)) {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    internal static void AddStandardNamespaces(XElement root, bool includePresentation = true) {
        root.SetAttributeValue(XNamespace.Xmlns + "office", OdfNamespaces.Office.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "text", OdfNamespaces.Text.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "table", OdfNamespaces.Table.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "draw", OdfNamespaces.Draw.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "style", OdfNamespaces.Style.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "number", OdfNamespaces.Number.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "fo", OdfNamespaces.Fo.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "svg", OdfNamespaces.Svg.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "xlink", OdfNamespaces.XLink.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "meta", OdfNamespaces.Meta.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "dc", OdfNamespaces.Dc.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "config", OdfNamespaces.Config.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "of", OdfNamespaces.Of.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "anim", OdfNamespaces.Anim.NamespaceName);
        root.SetAttributeValue(XNamespace.Xmlns + "smil", OdfNamespaces.Smil.NamespaceName);
        if (includePresentation) {
            root.SetAttributeValue(XNamespace.Xmlns + "presentation", OdfNamespaces.Presentation.NamespaceName);
        }
    }
}
