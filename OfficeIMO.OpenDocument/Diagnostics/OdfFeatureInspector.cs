namespace OfficeIMO.OpenDocument;

internal static class OdfFeatureInspector {
    private static readonly HashSet<string> KnownNamespaces = new HashSet<string>(StringComparer.Ordinal) {
        OdfNamespaces.Office.NamespaceName, OdfNamespaces.Text.NamespaceName, OdfNamespaces.Table.NamespaceName,
        OdfNamespaces.Draw.NamespaceName, OdfNamespaces.Presentation.NamespaceName, OdfNamespaces.Style.NamespaceName,
        OdfNamespaces.Number.NamespaceName, OdfNamespaces.Fo.NamespaceName, OdfNamespaces.Svg.NamespaceName,
        OdfNamespaces.XLink.NamespaceName, OdfNamespaces.Meta.NamespaceName, OdfNamespaces.Dc.NamespaceName,
        OdfNamespaces.Manifest.NamespaceName, OdfNamespaces.Config.NamespaceName, OdfNamespaces.Of.NamespaceName,
        XNamespace.Xml.NamespaceName, XNamespace.Xmlns.NamespaceName, string.Empty
    };

    internal static OdfFeatureReport Inspect(OdfPackage package) {
        var findings = new List<OdfFeatureFinding>();
        foreach (OdfPackageEntry entry in package.Entries.Where(entry => entry.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))) {
            XDocument document;
            try { document = package.GetXml(entry.Name); } catch { continue; }
            if (document.Root == null) continue;

            AddElementFinding(document, OdfNamespaces.Office + "scripts", "scripts", OdfFeatureSupport.Preserved, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Office + "annotation", "annotations", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Text + "tracked-changes", "tracked-changes", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Draw + "object", "embedded-objects", OdfFeatureSupport.Preserved, entry.Name, findings);

            var foreign = document.Root.DescendantsAndSelf()
                .Select(element => element.Name.NamespaceName)
                .Concat(document.Root.DescendantsAndSelf().Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).Select(attribute => attribute.Name.NamespaceName))
                .Where(namespaceName => !KnownNamespaces.Contains(namespaceName))
                .GroupBy(namespaceName => namespaceName, StringComparer.Ordinal);
            foreach (IGrouping<string, string> group in foreign) {
                findings.Add(new OdfFeatureFinding("foreign-namespace:" + group.Key, OdfFeatureSupport.Preserved, entry.Name, group.Count()));
            }
        }
        if (package.IsSigned) findings.Add(new OdfFeatureFinding("digital-signatures", OdfFeatureSupport.Preserved, "META-INF"));
        return new OdfFeatureReport(findings);
    }

    private static void AddElementFinding(XDocument document, XName elementName, string featureName, OdfFeatureSupport support,
        string partPath, List<OdfFeatureFinding> findings) {
        int count = document.Descendants(elementName).Count();
        if (count > 0) findings.Add(new OdfFeatureFinding(featureName, support, partPath, count));
    }
}
