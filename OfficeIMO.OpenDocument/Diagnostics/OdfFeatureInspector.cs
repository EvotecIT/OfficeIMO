namespace OfficeIMO.OpenDocument;

internal static class OdfFeatureInspector {
    private static readonly HashSet<string> KnownNamespaces = new HashSet<string>(StringComparer.Ordinal) {
        OdfNamespaces.Office.NamespaceName, OdfNamespaces.Text.NamespaceName, OdfNamespaces.Table.NamespaceName,
        OdfNamespaces.Draw.NamespaceName, OdfNamespaces.Presentation.NamespaceName, OdfNamespaces.Style.NamespaceName,
        OdfNamespaces.Number.NamespaceName, OdfNamespaces.Fo.NamespaceName, OdfNamespaces.Svg.NamespaceName,
        OdfNamespaces.XLink.NamespaceName, OdfNamespaces.Meta.NamespaceName, OdfNamespaces.Dc.NamespaceName,
        OdfNamespaces.Manifest.NamespaceName, OdfNamespaces.Config.NamespaceName, OdfNamespaces.Of.NamespaceName,
        OdfNamespaces.Anim.NamespaceName, OdfNamespaces.Smil.NamespaceName, OdfNamespaces.Script.NamespaceName,
        XNamespace.Xml.NamespaceName, XNamespace.Xmlns.NamespaceName, string.Empty
    };

    internal static OdfFeatureReport Inspect(OdfPackage package) {
        var findings = new List<OdfFeatureFinding>();
        var diagnostics = new List<OdfFeatureDiagnostic>();
        foreach (OdfPackageEntry entry in package.Entries.Where(entry => entry.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))) {
            XDocument document;
            try {
                document = package.GetXml(entry.Name);
            } catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
                diagnostics.Add(new OdfFeatureDiagnostic(
                    "ODF_FEATURE_XML_UNREADABLE",
                    entry.Name,
                    "The XML part could not be parsed and its features were not classified: " + exception.Message));
                continue;
            }
            if (document.Root == null) continue;

            int scripts = document.Descendants(OdfNamespaces.Office + "scripts")
                .Sum(element => element.Elements().Count());
            if (scripts > 0) findings.Add(new OdfFeatureFinding(
                "scripts", OdfFeatureSupport.Preserved, entry.Name, scripts));
            AddElementFinding(document, OdfNamespaces.Office + "annotation", "annotations", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Text + "tracked-changes", "tracked-changes", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Draw + "object", "embedded-objects", OdfFeatureSupport.Preserved, entry.Name, findings);
            int eventListeners = document.Descendants().Count(element => element.Name.LocalName == "event-listener");
            if (eventListeners > 0) findings.Add(new OdfFeatureFinding("event-listeners", OdfFeatureSupport.Preserved, entry.Name, eventListeners));
            int externalLinks = document.Root.DescendantsAndSelf().Attributes(OdfNamespaces.XLink + "href")
                .Count(attribute => IsExternalHref(attribute.Value));
            if (externalLinks > 0) findings.Add(new OdfFeatureFinding("external-links", OdfFeatureSupport.Preserved, entry.Name, externalLinks));
            AddElementFinding(document, OdfNamespaces.Table + "content-validation", "spreadsheet-validations", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "database-range", "spreadsheet-database-ranges", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "filter", "spreadsheet-filters", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "named-range", "spreadsheet-named-ranges", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "named-expression", "spreadsheet-named-expressions", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "scenario", "spreadsheet-scenarios", OdfFeatureSupport.Preserved, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "detective", "spreadsheet-detective", OdfFeatureSupport.Preserved, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Text + "note", "text-notes", OdfFeatureSupport.Inspected, entry.Name, findings);
            int bookmarks = document.Descendants(OdfNamespaces.Text + "bookmark").Count()
                + document.Descendants(OdfNamespaces.Text + "bookmark-start").Count();
            if (bookmarks > 0) findings.Add(new OdfFeatureFinding("text-bookmarks", OdfFeatureSupport.Editable, entry.Name, bookmarks));
            AddElementFinding(document, OdfNamespaces.Text + "section", "text-sections", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Text + "table-of-content", "text-tables-of-content", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Presentation + "notes", "presentation-notes", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Style + "master-page", "master-pages", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Draw + "plugin", "embedded-media", OdfFeatureSupport.Preserved, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Draw + "object-ole", "embedded-ole-objects", OdfFeatureSupport.Preserved, entry.Name, findings);
            int formulas = document.Descendants(OdfNamespaces.Table + "table-cell")
                .Count(element => element.Attribute(OdfNamespaces.Table + "formula") != null);
            if (formulas > 0) findings.Add(new OdfFeatureFinding("spreadsheet-formulas", OdfFeatureSupport.Editable, entry.Name, formulas));
            int transitions = document.Descendants(OdfNamespaces.Style + "drawing-page-properties")
                .Count(element => element.Attribute(OdfNamespaces.Presentation + "transition-type") != null || element.Attribute(OdfNamespaces.Presentation + "transition-style") != null);
            if (transitions > 0) findings.Add(new OdfFeatureFinding("presentation-transitions", OdfFeatureSupport.Editable, entry.Name, transitions));
            int animations = document.Descendants(OdfNamespaces.Anim + "animate").Count();
            if (animations > 0) findings.Add(new OdfFeatureFinding("presentation-animations", OdfFeatureSupport.Editable, entry.Name, animations));

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
        return new OdfFeatureReport(findings, diagnostics);
    }

    private static void AddElementFinding(XDocument document, XName elementName, string featureName, OdfFeatureSupport support,
        string partPath, List<OdfFeatureFinding> findings) {
        int count = document.Descendants(elementName).Count();
        if (count > 0) findings.Add(new OdfFeatureFinding(featureName, support, partPath, count));
    }

    private static bool IsExternalHref(string href) {
        if (string.IsNullOrWhiteSpace(href) || href.StartsWith("#", StringComparison.Ordinal)) return false;
        return href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _);
    }
}
