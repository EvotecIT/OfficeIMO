using OfficeIMO.Core.Internal;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    private static readonly XNamespace OfficeOoo = "http://openoffice.org/2009/office";
    private const string MetaPart = "meta.xml";
    private const string ContentPart = "content.xml";
    private const string StylesPart = "styles.xml";
    private const string SettingsPart = "settings.xml";
    private const string ThumbnailEntry = "Thumbnails/thumbnail.png";
    private const string ConditionalNote = "Shown or hidden by a field condition; evaluate the condition and remove the content in the document.";

    /// <summary>Finds metadata, personal information and hidden data without changing the document.</summary>
    public OfficeHiddenDataReport InspectHiddenData() => CollectHiddenData(null);

    /// <summary>Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item. Save the document to keep the result.</summary>
    /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps tracked changes, hidden content and provenance.</param>
    public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) =>
        CollectHiddenData(options ?? new OfficeHiddenDataOptions());

    private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
        var collector = new OfficeHiddenDataCollector(Kind switch {
            OdfDocumentKind.Text => "ODT",
            OdfDocumentKind.Spreadsheet => "ODS",
            OdfDocumentKind.Presentation => "ODP",
            _ => "ODG"
        }, removal);

        CollectOdfDocumentProperties(collector);
        CollectOdfCustomProperties(collector);
        CollectOdfComments(collector);
        CollectOdfTrackedChanges(collector);
        CollectOdfIdentities(collector);
        CollectOdfHiddenContent(collector);
        CollectOdfSpeakerNotes(collector);
        CollectOdfThumbnail(collector);
        CollectOdfImagesAndProvenance(collector);
        CollectOdfPrinterSettings(collector);
        CollectOdfExternalReferences(collector);
        return collector.Build();
    }

    private void CollectOdfDocumentProperties(OfficeHiddenDataCollector collector) {
        XElement? meta = GetExistingXml(MetaPart)?.Root?.Element(OdfNamespaces.Office + "meta");
        if (meta == null) return;
        AddMeta(OdfNamespaces.Meta + "initial-creator", "Author");
        AddMeta(OdfNamespaces.Dc + "creator", "Last modified by");
        AddMeta(OdfNamespaces.Meta + "printed-by", "Last printed by");
        AddMeta(OdfNamespaces.Meta + "print-date", "Last printed");
        AddMeta(OdfNamespaces.Meta + "editing-cycles", "Revision number");
        AddMeta(OdfNamespaces.Meta + "editing-duration", "Total editing time");
        // LibreOffice-style generators carry the build and operating system after '$'; a plain product name identifies nobody.
        foreach (XElement generator in meta.Elements(OdfNamespaces.Meta + "generator").Where(element => element.Value.IndexOf('$') >= 0).ToList()) {
            collector.Add(OfficeHiddenDataKind.DocumentProperties, MetaPart, "Generator with build and platform", generator.Value,
                () => RemoveXml(generator, MetaPart));
        }
        foreach (XElement template in meta.Elements(OdfNamespaces.Meta + "template").ToList()) {
            collector.Add(OfficeHiddenDataKind.DocumentProperties, MetaPart, "Template",
                (string?)template.Attribute(OdfNamespaces.XLink + "href"), () => RemoveXml(template, MetaPart));
        }

        void AddMeta(XName name, string description) {
            foreach (XElement element in meta.Elements(name).ToList()) {
                collector.Add(OfficeHiddenDataKind.DocumentProperties, MetaPart, description, element.Value, () => RemoveXml(element, MetaPart));
            }
        }
    }

    private void CollectOdfCustomProperties(OfficeHiddenDataCollector collector) {
        XElement? meta = GetExistingXml(MetaPart)?.Root?.Element(OdfNamespaces.Office + "meta");
        foreach (XElement property in meta?.Elements(OdfNamespaces.Meta + "user-defined").ToList() ?? new List<XElement>()) {
            collector.Add(OfficeHiddenDataKind.CustomProperties, MetaPart,
                "Custom property '" + (string?)property.Attribute(OdfNamespaces.Meta + "name") + "'", property.Value,
                () => RemoveXml(property, MetaPart));
        }
    }

    private void CollectOdfComments(OfficeHiddenDataCollector collector) {
        foreach ((string part, XDocument document) in ContentParts()) {
            foreach (XElement annotation in Annotations(document).ToList()) {
                collector.Add(OfficeHiddenDataKind.Comments, Locate(annotation, part),
                    "Comment by " + (annotation.Element(OdfNamespaces.Dc + "creator")?.Value ?? "unknown author"),
                    AnnotationText(annotation), () => {
                        string? name = (string?)annotation.Attribute(OdfNamespaces.Office + "name");
                        if (name != null) {
                            document.Descendants(OdfNamespaces.Office + "annotation-end")
                                .Where(end => (string?)end.Attribute(OdfNamespaces.Office + "name") == name).Remove();
                        }
                        RemoveXml(annotation, part);
                    });
            }
        }
    }

    private void CollectOdfTrackedChanges(OfficeHiddenDataCollector collector) {
        if (GetExistingXml(ContentPart) is not XDocument content) return;
        OfficeTrackedChangesAction action = collector.Removal == null ? OfficeTrackedChangesAction.Keep : collector.TrackedChanges;
        string? keptNote = collector.Removal == null ? null : OfficeHiddenDataCollector.TrackedChangesKeptText;

        if (this is OdtDocument text) {
            foreach (XElement region in content.Descendants(OdfNamespaces.Text + "tracked-changes")
                         .Elements(OdfNamespaces.Text + "changed-region").ToList()) {
                XElement? change = region.Elements().FirstOrDefault(element => element.Name.Namespace == OdfNamespaces.Text);
                string kind = change?.Name.LocalName ?? "change";
                string id = (string?)region.Attribute(OdfNamespaces.Text + "id") ?? (string?)region.Attribute(XNamespace.Xml + "id") ?? string.Empty;
                string description = "Tracked " + kind + " by " + (change?.Element(OdfNamespaces.Office + "change-info")?.Element(OdfNamespaces.Dc + "creator")?.Value ?? "unknown author");
                string? value = kind == "deletion" ? OdfTextCodec.ReadJoined(change!.Elements().Where(element =>
                    element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")) : null;

                OfficeHiddenDataOutcome outcome;
                string? note = null;
                if (action == OfficeTrackedChangesAction.Keep) {
                    outcome = collector.Removal == null ? OfficeHiddenDataOutcome.Found : OfficeHiddenDataOutcome.Retained;
                    note = keptNote;
                } else if (kind == "format-change") {
                    // ODF does not record the previous formatting, so a format change can be accepted but not rejected.
                    if (action == OfficeTrackedChangesAction.AcceptAll) {
                        RemoveChangeMarkers(content, id);
                        RemoveXml(region, ContentPart);
                        outcome = OfficeHiddenDataOutcome.Accepted;
                    } else {
                        outcome = OfficeHiddenDataOutcome.Retained;
                        note = "OpenDocument does not record the formatting a format change replaced, so it cannot be rejected.";
                    }
                } else if (id.Length > 0 && (action == OfficeTrackedChangesAction.AcceptAll ? text.AcceptTrackedChange(id) : text.RejectTrackedChange(id))) {
                    outcome = action == OfficeTrackedChangesAction.AcceptAll ? OfficeHiddenDataOutcome.Accepted : OfficeHiddenDataOutcome.Rejected;
                } else {
                    outcome = OfficeHiddenDataOutcome.Retained;
                    note = "The tracked change has no identifier the document can resolve.";
                }
                collector.Record(OfficeHiddenDataKind.TrackedChanges, ContentPart, description, value, outcome, note);
            }
            return;
        }

        foreach (XElement change in content.Descendants(OdfNamespaces.Table + "tracked-changes").Elements().ToList()) {
            OfficeHiddenDataOutcome outcome = collector.Removal == null ? OfficeHiddenDataOutcome.Found : OfficeHiddenDataOutcome.Retained;
            string? note = collector.Removal == null ? null : action == OfficeTrackedChangesAction.Keep ? keptNote
                : "Spreadsheet tracked changes are reported only; accept or reject them in a spreadsheet application.";
            collector.Record(OfficeHiddenDataKind.TrackedChanges, ContentPart,
                "Tracked " + change.Name.LocalName + " by " + (change.Descendants(OdfNamespaces.Dc + "creator").FirstOrDefault()?.Value ?? "unknown author"),
                null, outcome, note);
        }
    }

    private static void RemoveChangeMarkers(XDocument content, string id) {
        content.Descendants().Where(element => element.Name.Namespace == OdfNamespaces.Text &&
                (element.Name.LocalName is "change-start" or "change-end" or "change") &&
                (string?)element.Attribute(OdfNamespaces.Text + "change-id") == id).Remove();
    }

    /// <summary>Authors on comments and tracked changes that stay in the document.</summary>
    private void CollectOdfIdentities(OfficeHiddenDataCollector collector) {
        var authored = new List<(string Name, string Location, string Part, XElement Owner)>();
        foreach ((string part, XDocument document) in ContentParts()) {
            foreach (XElement annotation in Annotations(document)) {
                XElement? creator = annotation.Element(OdfNamespaces.Dc + "creator");
                if (creator != null && creator.Value.Length > 0) authored.Add((creator.Value, Locate(annotation, part), part, annotation));
            }
            foreach (XElement info in document.Descendants(OdfNamespaces.Office + "change-info")) {
                XElement? creator = info.Element(OdfNamespaces.Dc + "creator");
                if (creator != null && creator.Value.Length > 0) authored.Add((creator.Value, "Tracked changes in " + part, part, info));
            }
        }

        foreach (IGrouping<string, (string Name, string Location, string Part, XElement Owner)> group in authored.GroupBy(item => item.Name, StringComparer.Ordinal)) {
            if (collector.Removal != null && group.Key == collector.Removal.AnonymousAuthor) continue;
            collector.Add(OfficeHiddenDataKind.PersonalIdentities, string.Join(", ", group.Select(item => item.Location).Distinct()),
                "Author on " + group.Count().ToString(CultureInfo.InvariantCulture) + " comments or tracked changes", group.Key, () => {
                    foreach ((_, _, string part, XElement owner) in group) {
                        owner.Element(OdfNamespaces.Dc + "creator")!.Value = collector.Removal!.AnonymousAuthor;
                        foreach (XElement initials in owner.Elements().Where(element => element.Name.LocalName is "creator-initials" or "sender-initials")) {
                            initials.Value = collector.Removal.AnonymousInitials;
                        }
                        MarkPartDirty(part);
                    }
                }, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
        }
    }

    private void CollectOdfHiddenContent(OfficeHiddenDataCollector collector) {
        const OfficeHiddenDataKind kind = OfficeHiddenDataKind.HiddenContent;
        if (GetExistingXml(ContentPart) is not XDocument content) return;
        switch (this) {
            case OdtDocument:
                foreach (XElement section in content.Descendants(OdfNamespaces.Text + "section").ToList()) {
                    string? display = (string?)section.Attribute(OdfNamespaces.Text + "display");
                    string location = "Section '" + (string?)section.Attribute(OdfNamespaces.Text + "name") + "'";
                    if (display == "none") {
                        collector.Add(kind, location, "Hidden section", OdfTextCodec.ReadJoined(section.Descendants(OdfNamespaces.Text + "p")),
                            () => RemoveXml(section, ContentPart));
                    } else if (display == "condition") {
                        collector.Add(kind, location, "Conditionally hidden section", (string?)section.Attribute(OdfNamespaces.Text + "condition"), null, ConditionalNote);
                    }
                }
                foreach (XElement field in content.Descendants().Where(element => element.Name == OdfNamespaces.Text + "hidden-text" ||
                             element.Name == OdfNamespaces.Text + "hidden-paragraph").ToList()) {
                    collector.Add(kind, ContentPart, field.Name.LocalName == "hidden-text" ? "Conditionally hidden text" : "Conditionally hidden paragraph",
                        (string?)field.Attribute(OdfNamespaces.Text + "string-value") ?? (string?)field.Attribute(OdfNamespaces.Text + "condition"), null, ConditionalNote);
                }
                HashSet<string> hiddenStyles = HiddenTextStyles();
                foreach (XElement element in content.Descendants().Where(element =>
                             (element.Name == OdfNamespaces.Text + "span" || element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h") &&
                             hiddenStyles.Contains((string?)element.Attribute(OdfNamespaces.Text + "style-name") ?? string.Empty)).ToList()) {
                    if (element.Ancestors().Any(ancestor => ancestor.Name == OdfNamespaces.Text + "tracked-changes") || element.Parent == null) continue;
                    if (element.Ancestors().Any(ancestor => hiddenStyles.Contains((string?)ancestor.Attribute(OdfNamespaces.Text + "style-name") ?? string.Empty))) continue;
                    collector.Add(kind, ContentPart, element.Name.LocalName == "span" ? "Hidden text" : "Hidden paragraph",
                        OdfTextCodec.ReadJoined(new[] { element }), () => RemoveXml(element, ContentPart));
                }
                break;
            case OdsDocument spreadsheet:
                HashSet<string> hiddenTableStyles = HiddenTableStyles(content);
                foreach (OdsSheet sheet in spreadsheet.Sheets) {
                    bool hidden = sheet.Hidden || hiddenTableStyles.Contains((string?)sheet.Element.Attribute(OdfNamespaces.Table + "style-name") ?? string.Empty);
                    string location = "Sheet '" + sheet.Name + "'";
                    if (hidden) {
                        bool last = spreadsheet.Sheets.Count == 1;
                        collector.Add(kind, location, "Hidden sheet", null, last ? null : () => spreadsheet.RemoveSheet(sheet.Name),
                            "A spreadsheet must keep at least one sheet.");
                        if (!last && collector.Removes(kind)) continue;
                    }
                    int rows = sheet.Element.Descendants(OdfNamespaces.Table + "table-row").Count(IsCollapsed);
                    int columns = sheet.Element.Descendants(OdfNamespaces.Table + "table-column").Count(IsCollapsed);
                    if (rows + columns > 0) {
                        collector.Add(kind, location, "Hidden rows or columns",
                            rows.ToString(CultureInfo.InvariantCulture) + " row and " + columns.ToString(CultureInfo.InvariantCulture) + " column declarations",
                            null, "Deleting rows or columns here would not update the formulas that refer to them; delete them in a spreadsheet application.");
                    }
                }
                break;
            case OdpPresentation presentation:
                foreach (OdpSlide slide in presentation.Slides) {
                    string location = "Slide '" + slide.Name + "'";
                    if (slide.Hidden) {
                        bool last = presentation.Slides.Count == 1;
                        collector.Add(kind, location, "Hidden slide", null, last ? null : () => RemoveXml(slide.Element, ContentPart),
                            "A presentation must keep at least one slide.");
                        if (!last && collector.Removes(kind)) continue;
                    }
                    foreach (OdpShape shape in slide.Shapes.Where(shape => shape.Hidden).ToList()) {
                        collector.Add(kind, location + ", shape '" + (string?)shape.Element.Attribute(OdfNamespaces.Draw + "name") + "'", "Hidden shape", null,
                            () => RemoveXml(shape.Element, ContentPart));
                    }
                }
                break;
            case OdgDocument drawing:
                foreach (OdgPage page in drawing.Pages) {
                    var hiddenLayers = new HashSet<string>(page.EffectiveLayers.Where(layer => !layer.IsVisible() && !layer.IsVisible(forPrint: true))
                        .Select(layer => layer.Name), StringComparer.Ordinal);
                    if (hiddenLayers.Count == 0) continue;
                    foreach (XElement shape in page.Element.Elements().Where(element =>
                                 hiddenLayers.Contains((string?)element.Attribute(OdfNamespaces.Draw + "layer") ?? string.Empty)).ToList()) {
                        collector.Add(kind, "Page '" + page.Name + "', shape '" + (string?)shape.Attribute(OdfNamespaces.Draw + "name") + "'",
                            "Shape on hidden layer '" + (string?)shape.Attribute(OdfNamespaces.Draw + "layer") + "'", null,
                            () => RemoveXml(shape, ContentPart));
                    }
                }
                break;
        }

        static bool IsCollapsed(XElement element) => (string?)element.Attribute(OdfNamespaces.Table + "visibility") is "collapse" or "filter";
    }

    /// <summary>Text and paragraph styles that hide their text, declared on the style or inherited from a parent style.</summary>
    private HashSet<string> HiddenTextStyles() {
        var styles = new Dictionary<string, (string? Parent, string? Display)>(StringComparer.Ordinal);
        foreach ((_, XDocument document) in ContentParts()) {
            foreach (XElement style in document.Descendants(OdfNamespaces.Style + "style").Where(style =>
                         (string?)style.Attribute(OdfNamespaces.Style + "family") is "text" or "paragraph")) {
                string? name = (string?)style.Attribute(OdfNamespaces.Style + "name");
                if (name == null || styles.ContainsKey(name)) continue;
                styles[name] = ((string?)style.Attribute(OdfNamespaces.Style + "parent-style-name"),
                    (string?)style.Element(OdfNamespaces.Style + "text-properties")?.Attribute(OdfNamespaces.Text + "display"));
            }
        }
        var hidden = new HashSet<string>(StringComparer.Ordinal);
        foreach (string name in styles.Keys) {
            string? current = name;
            for (int depth = 0; current != null && depth < 64 && styles.TryGetValue(current, out var style); depth++) {
                if (style.Display != null) {
                    if (style.Display == "none") hidden.Add(name);
                    break;
                }
                current = style.Parent;
            }
        }
        return hidden;
    }

    private static HashSet<string> HiddenTableStyles(XDocument content) => new(content.Descendants(OdfNamespaces.Style + "style")
        .Where(style => (string?)style.Attribute(OdfNamespaces.Style + "family") == "table" &&
            (string?)style.Element(OdfNamespaces.Style + "table-properties")?.Attribute(OdfNamespaces.Table + "display") == "false")
        .Select(style => (string?)style.Attribute(OdfNamespaces.Style + "name") ?? string.Empty), StringComparer.Ordinal);

    private void CollectOdfSpeakerNotes(OfficeHiddenDataCollector collector) {
        if (this is not OdpPresentation presentation) return;
        foreach (OdpSlide slide in presentation.Slides) {
            foreach (XElement notes in slide.Element.Elements(OdfNamespaces.Presentation + "notes").ToList()) {
                collector.Add(OfficeHiddenDataKind.SpeakerNotes, "Slide '" + slide.Name + "'", "Speaker notes",
                    OdfTextCodec.ReadJoined(notes.Descendants(OdfNamespaces.Text + "p")), () => RemoveXml(notes, ContentPart));
            }
        }
    }

    private void CollectOdfThumbnail(OfficeHiddenDataCollector collector) {
        if (!Package.ContainsEntry(ThumbnailEntry)) return;
        collector.Add(OfficeHiddenDataKind.Thumbnail, ThumbnailEntry, "Document thumbnail image", null, () => Package.RemoveEntry(ThumbnailEntry));
    }

    /// <summary>Provenance and image metadata in the package and its images, through the shared asset engine.</summary>
    private void CollectOdfImagesAndProvenance(OfficeHiddenDataCollector collector) {
        if (Package.ContainsEntry(OfficeHiddenDataAssets.PackageManifestPath)) {
            byte[] manifest = Package.GetRequiredEntry(OfficeHiddenDataAssets.PackageManifestPath).GetBytesForSave();
            if (OfficeHiddenDataAssets.ProcessPackageManifest(manifest, OfficeHiddenDataAssets.PackageManifestPath, collector)) {
                // Saving rebuilds META-INF/manifest.xml from the remaining entries, which drops the file entry.
                Package.RemoveEntry(OfficeHiddenDataAssets.PackageManifestPath);
            }
        }

        foreach (OdfPackageEntry entry in Package.Entries.Where(entry => OdfPackage.IsBinaryPayloadPath(entry.Name) &&
                     !entry.Name.StartsWith("Thumbnails/", StringComparison.Ordinal) && !entry.Name.StartsWith("META-INF/", StringComparison.Ordinal) &&
                     (entry.Name.StartsWith("Pictures/", StringComparison.Ordinal) || entry.Name.Contains("/Pictures/") ||
                      (entry.MediaType ?? string.Empty).StartsWith("image/", StringComparison.OrdinalIgnoreCase))).ToList()) {
            byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(entry.GetBytesForSave(), entry.Name, collector);
            if (replacement != null) Package.AddOrReplaceEntry(entry.Name, replacement, entry.MediaType ?? string.Empty);
        }
    }

    private void CollectOdfPrinterSettings(OfficeHiddenDataCollector collector) {
        if (GetExistingXml(SettingsPart) is not XDocument settings) return;
        foreach (XElement item in settings.Descendants(OdfNamespaces.Config + "config-item").Where(item =>
                     (string?)item.Attribute(OdfNamespaces.Config + "name") is "PrinterName" or "PrinterSetup").ToList()) {
            if (item.Value.Length == 0) continue;
            string name = (string?)item.Attribute(OdfNamespaces.Config + "name") ?? string.Empty;
            collector.Add(OfficeHiddenDataKind.PrinterSettings, SettingsPart, name == "PrinterName" ? "Printer name" : "Printer driver setup",
                name == "PrinterName" ? item.Value : null, () => RemoveXml(item, SettingsPart));
        }
    }

    /// <summary>Linked files and sources. They are the content the document shows, so they are reported and kept.</summary>
    private void CollectOdfExternalReferences(OfficeHiddenDataCollector collector) {
        foreach ((string part, XDocument document) in ContentParts()) {
            foreach (XElement element in document.Descendants().Where(element =>
                         element.Name == OdfNamespaces.Draw + "image" || element.Name == OdfNamespaces.Draw + "object" ||
                         element.Name == OdfNamespaces.Draw + "object-ole" || element.Name == OdfNamespaces.Text + "section-source" ||
                         element.Name == OdfNamespaces.Table + "table-source" || element.Name == OdfNamespaces.Table + "cell-range-source")) {
                string? href = (string?)element.Attribute(OdfNamespaces.XLink + "href");
                if (href == null || !IsExternalHref(href)) continue;
                collector.Add(OfficeHiddenDataKind.ExternalReferences, Locate(element, part), "Linked " + element.Name.LocalName.Replace('-', ' '),
                    href, null, "The document shows this linked content; removing the link would remove the content.");
            }
        }
    }

    private static bool IsExternalHref(string href) {
        string value = href.Trim();
        if (value.Length == 0 || value.StartsWith("#", StringComparison.Ordinal)) return false;
        if (value.StartsWith("./", StringComparison.Ordinal)) return false;
        return value.StartsWith("../", StringComparison.Ordinal) || value.StartsWith("/", StringComparison.Ordinal) ||
            value.StartsWith("\\", StringComparison.Ordinal) || value.IndexOf(':') > 0;
    }

    private IEnumerable<(string Part, XDocument Document)> ContentParts() {
        foreach (string part in new[] { ContentPart, StylesPart }) {
            if (GetExistingXml(part) is XDocument document) yield return (part, document);
        }
    }

    private static IEnumerable<XElement> Annotations(XDocument document) => document.Descendants().Where(element =>
        element.Name == OdfNamespaces.Office + "annotation" || element.Name == OfficeOoo + "annotation");

    private static string AnnotationText(XElement annotation) => OdfTextCodec.ReadJoined(annotation.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h"));

    private string Locate(XElement element, string part) {
        XElement? table = Kind == OdfDocumentKind.Spreadsheet ? element.Ancestors(OdfNamespaces.Table + "table").LastOrDefault() : null;
        if (table != null) return "Sheet '" + (string?)table.Attribute(OdfNamespaces.Table + "name") + "'";
        XElement? page = element.Ancestors(OdfNamespaces.Draw + "page").FirstOrDefault();
        if (page != null) return (Kind == OdfDocumentKind.Presentation ? "Slide '" : "Page '") + (string?)page.Attribute(OdfNamespaces.Draw + "name") + "'";
        return part;
    }

    private void RemoveXml(XElement element, string part) {
        if (element.Parent == null) return;
        element.Remove();
        MarkPartDirty(part);
    }
}
