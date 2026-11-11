using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private const string BibliographyNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
    private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /// <summary>Finds metadata, personal information and hidden data, the way Office's Document Inspector does, without changing the document.</summary>
    public OfficeHiddenDataReport InspectHiddenData() => CollectHiddenData(null);

    /// <summary>Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item. Save the document to keep the result.</summary>
    /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps tracked changes and hidden content.</param>
    public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) =>
        CollectHiddenData(options ?? new OfficeHiddenDataOptions());

    private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart
            ?? throw new InvalidOperationException("MainDocumentPart is missing.");
        var collector = new OfficeHiddenDataCollector("Word", removal);

        CollectWordTrackedChanges(collector);
        CollectWordComments(main, collector);
        CollectWordHiddenText(main, collector);
        CollectWordIdentities(main, collector);
        OfficeOpenXmlHiddenData.InspectDocumentProperties(_wordprocessingDocument, collector);
        OfficeOpenXmlHiddenData.InspectCustomProperties(_wordprocessingDocument, collector, name => CustomProperties.Remove(name));
        CollectWordSettings(main, collector);
        CollectWordRevisionIdentifiers(main, collector);
        CollectWordExternalReferences(collector);
        OfficeOpenXmlHiddenData.InspectCustomXml(main, main.CustomXmlParts, collector, part => GetRetainedCustomXmlReason(main, part));
        OfficeOpenXmlHiddenData.InspectThumbnail(_wordprocessingDocument, collector);
        OfficeOpenXmlHiddenData.InspectImageMetadata(_wordprocessingDocument, collector);
        OfficeOpenXmlHiddenData.InspectPrinterSettings(_wordprocessingDocument, collector);
        return collector.Build();
    }

    private void CollectWordTrackedChanges(OfficeHiddenDataCollector collector) {
        IReadOnlyList<WordRevisionInfo> revisions = InspectReview().Revisions;
        if (revisions.Count == 0) return;
        OfficeTrackedChangesAction action = collector.Removal == null ? OfficeTrackedChangesAction.Keep : collector.TrackedChanges;
        if (action == OfficeTrackedChangesAction.AcceptAll) AcceptRevisions();
        else if (action == OfficeTrackedChangesAction.RejectAll) RejectRevisions();

        OfficeHiddenDataOutcome outcome = action switch {
            OfficeTrackedChangesAction.AcceptAll => OfficeHiddenDataOutcome.Accepted,
            OfficeTrackedChangesAction.RejectAll => OfficeHiddenDataOutcome.Rejected,
            _ => collector.Removal == null ? OfficeHiddenDataOutcome.Found : OfficeHiddenDataOutcome.Retained
        };
        string? note = outcome == OfficeHiddenDataOutcome.Retained ? "Tracked changes are kept unless TrackedChanges is AcceptAll or RejectAll." : null;
        foreach (WordRevisionInfo revision in revisions) {
            string text = revision.AffectedText.Length > 0 ? revision.AffectedText : revision.LocationText;
            collector.Record(OfficeHiddenDataKind.TrackedChanges, revision.LocationKind + " " + revision.PartUri,
                revision.RevisionType + " by " + (revision.Author ?? "unknown author"), text, outcome, note);
        }
    }

    private void CollectWordComments(MainDocumentPart main, OfficeHiddenDataCollector collector) {
        foreach (WordComment comment in WordComment.GetAllComments(this)) {
            collector.Add(OfficeHiddenDataKind.Comments, "Comment " + comment.Id,
                "Comment by " + (comment.Author ?? "unknown author"), comment.Text, comment.Remove);
        }

        if (collector.Removes(OfficeHiddenDataKind.Comments) && main.WordprocessingCommentsPart?.Comments?.HasChildren != true) {
            if (main.WordprocessingCommentsPart != null) main.DeletePart(main.WordprocessingCommentsPart);
            if (main.WordprocessingCommentsExPart != null) main.DeletePart(main.WordprocessingCommentsExPart);
            if (main.WordprocessingCommentsIdsPart != null) main.DeletePart(main.WordprocessingCommentsIdsPart);
            if (main.WordCommentsExtensiblePart != null) main.DeletePart(main.WordCommentsExtensiblePart);
        }
    }

    private static void CollectWordHiddenText(MainDocumentPart main, OfficeHiddenDataCollector collector) {
        var resolver = new WordStyleResolver(main);
        foreach ((string story, OpenXmlPartRootElement root) in GetHiddenDataStories(main, includeComments: false)) {
            int paragraphIndex = 0;
            foreach (W.Paragraph paragraph in root.Descendants<W.Paragraph>().ToList()) {
                paragraphIndex++;
                List<W.Run> hidden = paragraph.Elements<W.Run>()
                    .Where(run => !run.Descendants<W.FieldChar>().Any() && !run.Descendants<W.FieldCode>().Any() && resolver.Resolve(run).Vanish)
                    .ToList();
                if (hidden.Count == 0) continue;
                string text = string.Concat(hidden.SelectMany(run => run.Elements<W.Text>()).Select(item => item.Text));
                collector.Add(OfficeHiddenDataKind.HiddenContent,
                    story + " paragraph " + paragraphIndex.ToString(CultureInfo.InvariantCulture), "Hidden text", text,
                    () => hidden.ForEach(run => run.Remove()));
            }
        }
    }

    private static void CollectWordIdentities(MainDocumentPart main, OfficeHiddenDataCollector collector) {
        WordprocessingPeoplePart? people = main.WordprocessingPeoplePart;
        if (people != null) {
            string names = string.Join(", ", people.People?.Elements<DocumentFormat.OpenXml.Office2013.Word.Person>()
                .Select(person => person.Author?.Value).Where(name => !string.IsNullOrEmpty(name)) ?? Array.Empty<string>());
            collector.Add(OfficeHiddenDataKind.PersonalIdentities, people.Uri.ToString(),
                "People part with reviewer identities", names, () => main.DeletePart(people));
        }

        var stories = GetHiddenDataStories(main, includeComments: true).ToList();
        if (main.StyleDefinitionsPart?.Styles != null) stories.Add(("Styles", main.StyleDefinitionsPart.Styles));
        if (main.NumberingDefinitionsPart?.Numbering != null) stories.Add(("Numbering", main.NumberingDefinitionsPart.Numbering));

        var byAuthor = new Dictionary<string, List<(string Story, OpenXmlElement Element)>>(StringComparer.Ordinal);
        foreach ((string story, OpenXmlPartRootElement root) in stories) {
            foreach (OpenXmlElement element in root.Descendants()) {
                string? author = GetWordAttributeValue(element, "author");
                if (author == null) continue;
                if (!byAuthor.TryGetValue(author, out var elements)) byAuthor[author] = elements = new();
                elements.Add((story, element));
            }
        }

        foreach (KeyValuePair<string, List<(string Story, OpenXmlElement Element)>> group in byAuthor) {
            string anonymous = collector.Removal?.AnonymousAuthor ?? string.Empty;
            if (collector.Removal != null && group.Key == anonymous) continue;
            string locations = string.Join(", ", group.Value.Select(item => item.Story).Distinct());
            collector.Add(OfficeHiddenDataKind.PersonalIdentities, locations,
                "Author name on " + group.Value.Count.ToString(CultureInfo.InvariantCulture) + " comment or revision marks", group.Key,
                () => {
                    string initials = collector.Removal!.AnonymousInitials;
                    foreach ((_, OpenXmlElement element) in group.Value) {
                        element.SetAttribute(new OpenXmlAttribute("w", "author", WordprocessingNamespace, anonymous));
                        if (GetWordAttributeValue(element, "initials") != null) {
                            element.SetAttribute(new OpenXmlAttribute("w", "initials", WordprocessingNamespace, initials));
                        }
                    }
                }, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
        }
    }

    private void CollectWordSettings(MainDocumentPart main, OfficeHiddenDataCollector collector) {
        DocumentSettingsPart? settingsPart = main.DocumentSettingsPart;
        W.Settings? settings = settingsPart?.Settings;
        if (settingsPart == null || settings == null) return;
        string location = settingsPart.Uri.ToString();

        W.DocumentVariables? variables = settings.GetFirstChild<W.DocumentVariables>();
        foreach (W.DocumentVariable variable in variables?.Elements<W.DocumentVariable>().ToList() ?? new List<W.DocumentVariable>()) {
            string name = variable.Name?.Value ?? string.Empty;
            collector.Add(OfficeHiddenDataKind.DocumentVariables, location, "Document variable '" + name + "'", variable.Val?.Value, () => {
                variable.Remove();
                DocumentVariables.Remove(name);
            });
        }
        if (variables != null && collector.Removes(OfficeHiddenDataKind.DocumentVariables) && !variables.HasChildren) variables.Remove();

        W.AttachedTemplate? template = settings.GetFirstChild<W.AttachedTemplate>();
        string? templateId = template?.Id?.Value;
        if (template != null) {
            string? path = settingsPart.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == templateId)?.Uri.OriginalString;
            collector.Add(OfficeHiddenDataKind.ExternalReferences, location, "Attached template", path, () => {
                template.Remove();
                if (templateId != null && settingsPart.ExternalRelationships.Any(relationship => relationship.Id == templateId)) {
                    settingsPart.DeleteExternalRelationship(templateId);
                }
            });
        }
    }

    private static void CollectWordRevisionIdentifiers(MainDocumentPart main, OfficeHiddenDataCollector collector) {
        var roots = GetHiddenDataStories(main, includeComments: true).ToList();
        if (main.StyleDefinitionsPart?.Styles != null) roots.Add(("Styles", main.StyleDefinitionsPart.Styles));
        if (main.NumberingDefinitionsPart?.Numbering != null) roots.Add(("Numbering", main.NumberingDefinitionsPart.Numbering));
        if (main.DocumentSettingsPart?.Settings != null) roots.Add(("Settings", main.DocumentSettingsPart.Settings));

        foreach ((string story, OpenXmlPartRootElement root) in roots) {
            var attributes = new List<(OpenXmlElement Element, string LocalName)>();
            var elements = new List<OpenXmlElement>();
            foreach (OpenXmlElement element in root.Descendants()) {
                if (element.Parent is W.Rsids) continue;
                if (!element.HasAttributes && element is not W.Rsids && element is not W.Rsid) continue;
                if (element is W.Rsids || element is W.Rsid) {
                    elements.Add(element);
                    continue;
                }
                if (!element.HasAttributes) continue;
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (attribute.NamespaceUri == WordprocessingNamespace && attribute.LocalName.StartsWith("rsid", StringComparison.Ordinal)) {
                        attributes.Add((element, attribute.LocalName));
                    }
                }
            }
            if (attributes.Count == 0 && elements.Count == 0) continue;
            collector.Add(OfficeHiddenDataKind.RevisionIdentifiers, story,
                (attributes.Count + elements.Count).ToString(CultureInfo.InvariantCulture) + " revision save identifiers", null, () => {
                    foreach ((OpenXmlElement element, string localName) in attributes) element.RemoveAttribute(localName, WordprocessingNamespace);
                    foreach (OpenXmlElement element in elements) element.Remove();
                });
        }
    }

    private void CollectWordExternalReferences(OfficeHiddenDataCollector collector) {
        string? templateId = _wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<W.AttachedTemplate>()?.Id?.Value;
        foreach (OfficeOpenXmlHiddenData.PartPair pair in OfficeOpenXmlHiddenData.EnumerateParts(_wordprocessingDocument)
                     .GroupBy(pair => pair.Part).Select(group => group.First())) {
            foreach (ExternalRelationship relationship in pair.Part.ExternalRelationships) {
                if (relationship.RelationshipType == HyperlinkRelationshipType) continue;
                if (pair.Part is DocumentSettingsPart && relationship.Id == templateId) continue;
                collector.Add(OfficeHiddenDataKind.ExternalReferences, pair.Part.Uri.ToString(),
                    "Linked external content (" + relationship.RelationshipType.Substring(relationship.RelationshipType.LastIndexOf('/') + 1) + ")",
                    relationship.Uri.OriginalString, null, "Removing linked content would remove it from the document.");
            }
        }
    }

    private static string? GetRetainedCustomXmlReason(MainDocumentPart main, CustomXmlPart part) {
        using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
            using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null });
            try {
                if (reader.MoveToContent() == System.Xml.XmlNodeType.Element && reader.NamespaceURI == BibliographyNamespace) {
                    return "Holds the bibliography sources that citations use.";
                }
            } catch (System.Xml.XmlException) {
                // Unreadable custom XML is not bibliography data.
            }
        }

        string? itemId = OfficeOpenXmlHiddenData.GetDataStoreItemId(part);
        if (itemId == null) return null;
        foreach ((_, OpenXmlPartRootElement root) in GetHiddenDataStories(main, includeComments: true)) {
            if (root.Descendants<W.DataBinding>().Any(binding => string.Equals(binding.StoreItemId?.Value, itemId, StringComparison.OrdinalIgnoreCase))) {
                return "Content controls are bound to this data.";
            }
        }
        return null;
    }

    private static string? GetWordAttributeValue(OpenXmlElement element, string localName) {
        if (!element.HasAttributes) return null;
        foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
            if (attribute.LocalName == localName && attribute.NamespaceUri == WordprocessingNamespace) return attribute.Value;
        }
        return null;
    }

    private static IEnumerable<(string Story, OpenXmlPartRootElement Root)> GetHiddenDataStories(MainDocumentPart main, bool includeComments) {
        if (main.Document != null) yield return ("Document", main.Document);
        int index = 0;
        foreach (HeaderPart header in main.HeaderParts) if (header.Header != null) yield return ("Header " + (++index).ToString(CultureInfo.InvariantCulture), header.Header);
        index = 0;
        foreach (FooterPart footer in main.FooterParts) if (footer.Footer != null) yield return ("Footer " + (++index).ToString(CultureInfo.InvariantCulture), footer.Footer);
        if (main.FootnotesPart?.Footnotes != null) yield return ("Footnotes", main.FootnotesPart.Footnotes);
        if (main.EndnotesPart?.Endnotes != null) yield return ("Endnotes", main.EndnotesPart.Endnotes);
        if (includeComments && main.WordprocessingCommentsPart?.Comments != null) yield return ("Comments", main.WordprocessingCommentsPart.Comments);
    }
}
