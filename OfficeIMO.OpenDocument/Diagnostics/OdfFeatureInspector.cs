namespace OfficeIMO.OpenDocument;

internal static class OdfFeatureInspector {
    private static readonly HashSet<string> KnownNamespaces = new HashSet<string>(StringComparer.Ordinal) {
        OdfNamespaces.Office.NamespaceName, OdfNamespaces.Text.NamespaceName, OdfNamespaces.Table.NamespaceName,
        OdfNamespaces.Draw.NamespaceName, OdfNamespaces.Chart.NamespaceName, OdfNamespaces.Presentation.NamespaceName, OdfNamespaces.Style.NamespaceName,
        OdfNamespaces.Number.NamespaceName, OdfNamespaces.Fo.NamespaceName, OdfNamespaces.Svg.NamespaceName,
        OdfNamespaces.XLink.NamespaceName, OdfNamespaces.Meta.NamespaceName, OdfNamespaces.Dc.NamespaceName,
        OdfNamespaces.Manifest.NamespaceName, OdfNamespaces.Config.NamespaceName, OdfNamespaces.Of.NamespaceName,
        OdfNamespaces.Anim.NamespaceName, OdfNamespaces.Smil.NamespaceName, OdfNamespaces.Script.NamespaceName,
        XNamespace.Xml.NamespaceName, XNamespace.Xmlns.NamespaceName, string.Empty
    };

    internal static OdfFeatureReport Inspect(OdfDocument source) {
        OdfPackage package = source.Package;
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
            int editableStyleMaps = CountEditableStyleMaps(document, entry.Name);
            if (editableStyleMaps > 0) findings.Add(new OdfFeatureFinding(
                "conditional-style-maps", OdfFeatureSupport.Editable, entry.Name, editableStyleMaps));
            int otherStyleMaps = document.Descendants(OdfNamespaces.Style + "map").Count() - editableStyleMaps;
            if (otherStyleMaps > 0) findings.Add(new OdfFeatureFinding(
                "unmodeled-style-maps", OdfFeatureSupport.Preserved, entry.Name, otherStyleMaps));
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
            XElement[] dataPilots = document.Descendants(OdfNamespaces.Table + "data-pilot-table").ToArray();
            int editableDataPilots = dataPilots.Count(pivot =>
                source is OdsDocument spreadsheet && OdsDataPilotTable.IsEditableElement(spreadsheet, pivot));
            if (editableDataPilots > 0) findings.Add(new OdfFeatureFinding(
                "spreadsheet-data-pilot-tables", OdfFeatureSupport.Editable, entry.Name, editableDataPilots));
            if (dataPilots.Length > editableDataPilots) findings.Add(new OdfFeatureFinding(
                "spreadsheet-data-pilot-tables", OdfFeatureSupport.Inspected, entry.Name,
                dataPilots.Length - editableDataPilots));
            AddElementFinding(document, OdfNamespaces.Table + "named-range", "spreadsheet-named-ranges", OdfFeatureSupport.Editable, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "named-expression", "spreadsheet-named-expressions", OdfFeatureSupport.Inspected, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "scenario", "spreadsheet-scenarios", OdfFeatureSupport.Preserved, entry.Name, findings);
            AddElementFinding(document, OdfNamespaces.Table + "detective", "spreadsheet-detective", OdfFeatureSupport.Preserved, entry.Name, findings);
            XElement[] textFields = document.Descendants()
                .Where(element => OdtField.TryGetKind(element.Name, out _)).ToArray();
            int basicTextFields = textFields.Count(element =>
                OdtField.IsBasicElement(element) && IsEditableOdtField(package.Kind, entry.Name, element));
            if (basicTextFields > 0) findings.Add(new OdfFeatureFinding(
                "text-fields", OdfFeatureSupport.Editable, entry.Name, basicTextFields));
            if (textFields.Length > basicTextFields) findings.Add(new OdfFeatureFinding(
                "text-fields", OdfFeatureSupport.Inspected, entry.Name, textFields.Length - basicTextFields));
            XElement[] notes = document.Descendants(OdfNamespaces.Text + "note").ToArray();
            int editableNotes = notes.Count(note =>
                !note.Ancestors(OdfNamespaces.Text + "tracked-changes").Any() &&
                ((string?)note.Attribute(OdfNamespaces.Text + "note-class") is "footnote" or "endnote") &&
                note.Element(OdfNamespaces.Text + "note-body") is XElement body &&
                body.Elements().All(child => child.Name == OdfNamespaces.Text + "p"));
            if (editableNotes > 0) findings.Add(new OdfFeatureFinding(
                "text-notes", OdfFeatureSupport.Editable, entry.Name, editableNotes));
            if (notes.Length > editableNotes) findings.Add(new OdfFeatureFinding(
                "text-notes", OdfFeatureSupport.Inspected, entry.Name, notes.Length - editableNotes));
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
        try {
            if (package.IsSigned) findings.Add(new OdfFeatureFinding("digital-signatures", OdfFeatureSupport.Preserved, "META-INF"));
        } catch (InvalidDataException exception) {
            diagnostics.Add(new OdfFeatureDiagnostic(
                "ODF_FEATURE_SIGNATURE_UNREADABLE",
                "META-INF",
                "A signature-like entry could not be classified: " + exception.Message));
        }
        return new OdfFeatureReport(findings, diagnostics);
    }

    private static void AddElementFinding(XDocument document, XName elementName, string featureName, OdfFeatureSupport support,
        string partPath, List<OdfFeatureFinding> findings) {
        int count = document.Descendants(elementName).Count();
        if (count > 0) findings.Add(new OdfFeatureFinding(featureName, support, partPath, count));
    }

    internal static bool IsEditableOdtField(OdfDocumentKind kind, string partPath, XElement element) {
        if (kind != OdfDocumentKind.Text ||
            element.Ancestors().Any(ancestor => ancestor.Name == OdfNamespaces.Text + "tracked-changes" ||
                ancestor.Name == OdfNamespaces.Text + "note" ||
                ancestor.Name == OdfNamespaces.Office + "annotation")) return false;
        XElement? paragraph = element.Ancestors().FirstOrDefault(ancestor =>
            ancestor.Name == OdfNamespaces.Text + "p" || ancestor.Name == OdfNamespaces.Text + "h");
        if (paragraph == null) return false;
        for (XElement? inline = element.Parent; inline != null && !ReferenceEquals(inline, paragraph);
            inline = inline.Parent) {
            if (inline.Name != OdfNamespaces.Text + "span" && inline.Name != OdfNamespaces.Text + "a")
                return false;
        }
        if (partPath == "styles.xml") {
            if (paragraph.Parent?.Name != OdfNamespaces.Style + "header" &&
                paragraph.Parent?.Name != OdfNamespaces.Style + "footer") return false;
            XElement? firstMaster = element.Document?.Root?
                .Element(OdfNamespaces.Office + "master-styles")?
                .Elements(OdfNamespaces.Style + "master-page").FirstOrDefault();
            return ReferenceEquals(paragraph.Parent.Parent, firstMaster);
        }
        if (partPath != "content.xml") return false;
        int tableDepth = 0;
        for (XElement? container = paragraph.Parent; container != null; container = container.Parent) {
            if (container.Name == OdfNamespaces.Office + "text") return true;
            if (container.Name == OdfNamespaces.Table + "table" && ++tableDepth > 1) return false;
            if (container.Name == OdfNamespaces.Table + "table-cell" &&
                !ReferenceEquals(paragraph.Parent, container)) return false;
            if (container.Name != OdfNamespaces.Text + "section" &&
                container.Name != OdfNamespaces.Text + "list" &&
                container.Name != OdfNamespaces.Text + "list-item" &&
                container.Name != OdfNamespaces.Text + "list-header" &&
                container.Name != OdfNamespaces.Table + "table" &&
                container.Name != OdfNamespaces.Table + "table-header-rows" &&
                container.Name != OdfNamespaces.Table + "table-row" &&
                container.Name != OdfNamespaces.Table + "table-cell") return false;
        }
        return false;
    }

    private static int CountEditableStyleMaps(XDocument document, string partPath) {
        XElement? root = document.Root;
        if (root == null) return 0;
        IEnumerable<XElement> containers = partPath == "content.xml"
            ? root.Elements(OdfNamespaces.Office + "automatic-styles")
            : partPath == "styles.xml"
                ? root.Elements(OdfNamespaces.Office + "styles")
                    .Concat(root.Elements(OdfNamespaces.Office + "automatic-styles"))
                : Enumerable.Empty<XElement>();
        return containers.SelectMany(container => container.Elements(OdfNamespaces.Style + "style"))
            .Where(style => OdfStyleRepository.TryParseFamily(
                (string?)style.Attribute(OdfNamespaces.Style + "family"), out _))
            .Sum(style => style.Elements(OdfNamespaces.Style + "map").Count());
    }

    private static bool IsExternalHref(string href) {
        if (string.IsNullOrWhiteSpace(href) || href.StartsWith("#", StringComparison.Ordinal)) return false;
        return href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _);
    }
}
