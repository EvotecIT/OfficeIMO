using System.Globalization;
using System.Text;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>
/// Clean a file before sharing: one report of everything that would leave with the file, grouped by category, and one action
/// that writes a copy without the selected categories. Every format goes through <see cref="OfficeDocument"/>'s hidden-data
/// front door, which reports Content Credentials and AI labels as its Provenance category; Word and PowerPoint files also get
/// the accessibility checker.
/// </summary>
internal static class ShareTool {
    internal const string PersonalGroup = "Personal and hidden data";
    internal const string AccessibilityGroup = "Accessibility";
    internal const string TableHeaders = "fix-table-headers";
    internal const string SourceLabels = "source-labels";

    internal static ToolResultDocument Run(ToolSession session, string action, ToolOptions options) {
        SelectedDocument file = session.Input();
        if (!OfficeDocument.SupportsHiddenData(file.Name)) {
            throw new NotSupportedException("Choose a Word, Excel or PowerPoint file, a PDF, or a JPEG, PNG or WebP image.");
        }
        return action switch {
            "inspect" => Inspect(session, file),
            "remove" => Remove(session, file, options),
            _ => throw new NotSupportedException($"The clean-before-sharing tool has no '{action}' action.")
        };
    }

    /// <summary>One checkbox on the page: which findings it covers, its title, group, and what selecting it does. Default categories start selected.</summary>
    private sealed record Category(string Id, string Title, string Explanation, bool Preselected, Func<OfficeHiddenDataFinding, bool> Covers,
        string Group = PersonalGroup, bool Selectable = true, OfficeHiddenDataKind Kind = OfficeHiddenDataKind.None);

    private static Category Kind(OfficeHiddenDataKind kind, string id, string title, string explanation, bool preselected) =>
        new(id, title, explanation, preselected, finding => finding.Kind == kind, Kind: kind);

    private static readonly Category[] HiddenData = [
        Kind(OfficeHiddenDataKind.DocumentProperties, "properties", "Document properties",
            "Author, last saved by, company, manager, template and editing time.", true),
        Kind(OfficeHiddenDataKind.CustomProperties, "custom-properties", "Custom properties",
            "Extra name and value pairs added by people or other apps.", true),
        Kind(OfficeHiddenDataKind.Comments, "comments", "Comments and notes",
            "Comments, replies and notes, with the text they point to left in place.", true),
        Kind(OfficeHiddenDataKind.PersonalIdentities, "people", "Names of people",
            "Names on comments and changes that stay in the file are replaced with \"Author\".", true),
        Kind(OfficeHiddenDataKind.TrackedChanges, "tracked-changes", "Tracked changes",
            "Selecting this accepts every change, as Accept All does in Word. Left unselected, changes stay for the reader to see.", false),
        Kind(OfficeHiddenDataKind.HiddenContent, "hidden-content", "Hidden content",
            "Hidden text, sheets, rows, columns, slides, shapes and PDF layers. Removing it changes what readers get, and deleting hidden rows or sheets can break formulas that use them.", false),
        Kind(OfficeHiddenDataKind.SpeakerNotes, "speaker-notes", "Speaker notes",
            "Notes under each slide.", true),
        Kind(OfficeHiddenDataKind.Thumbnail, "thumbnail", "Preview picture",
            "A small picture of the first page that file browsers show.", true),
        Kind(OfficeHiddenDataKind.CustomXml, "custom-xml", "Custom XML data",
            "Data stored by add-ins and document management systems. Bibliography sources and data bound to content controls stay.", true),
        Kind(OfficeHiddenDataKind.ImageMetadata, "image-metadata", "Photo details",
            "Camera, location (GPS) and editing details inside pictures. The pictures themselves don't change.", true),
        Kind(OfficeHiddenDataKind.PrinterSettings, "printer-settings", "Printer settings",
            "Printer names and setup saved with the file.", true),
        Kind(OfficeHiddenDataKind.ExternalReferences, "external-references", "Links to other files",
            "Paths to templates and linked workbooks. Links that formulas or refreshes need are reported and kept.", true),
        Kind(OfficeHiddenDataKind.DocumentVariables, "document-variables", "Document variables",
            "Values stored by macros and templates.", true),
        Kind(OfficeHiddenDataKind.RevisionIdentifiers, "revision-ids", "Editing session IDs",
            "Codes Word adds for each editing session, which can link copies of a file.", true),
        Kind(OfficeHiddenDataKind.ActiveContent, "active-content", "Scripts and actions",
            "JavaScript and actions that run or open something when the file is opened or clicked.", true),
        Kind(OfficeHiddenDataKind.Attachments, "attachments", "Attached files",
            "Files carried inside the document. Selecting this removes them; the reader may need them.", false),
        Kind(OfficeHiddenDataKind.Bookmarks, "bookmarks", "Bookmarks",
            "The bookmark list beside the pages. Their titles can show draft headings or internal names.", false)
    ];

    /// <summary>
    /// Provenance has one category in the hidden-data report and three choices on the page, one per carrier, because a visitor can
    /// keep a Content Credentials record and drop an AI label. Labels that don't declare AI are reported but never removed.
    /// </summary>
    private static readonly Category[] Provenance = [
        new(ProvenanceText.Manifests, "Content Credentials",
            "A record claiming how the content was made or edited (C2PA). Its signature is not verified here.", false,
            finding => Carrier(finding) == OfficeProvenanceCarrierKind.C2paManifest, ProvenanceText.Group),
        new(ProvenanceText.References, "Link to Content Credentials",
            "Points to a record stored outside the file.", false,
            finding => Carrier(finding) == OfficeProvenanceCarrierKind.C2paExternalManifest, ProvenanceText.Group),
        new(ProvenanceText.Declarations, "AI source label",
            "", false,
            finding => Carrier(finding) is { } carrier && carrier is not (OfficeProvenanceCarrierKind.C2paManifest or OfficeProvenanceCarrierKind.C2paExternalManifest) && DeclaresAi(finding),
            ProvenanceText.Group),
        new(SourceLabels, "Other source label",
            "", false,
            finding => Carrier(finding) is { } carrier && carrier is not (OfficeProvenanceCarrierKind.C2paManifest or OfficeProvenanceCarrierKind.C2paExternalManifest) && !DeclaresAi(finding),
            ProvenanceText.Group, Selectable: false)
    ];

    private static readonly Category[] Categories = [.. HiddenData, .. Provenance];

    private static OfficeProvenanceCarrierKind? Carrier(OfficeHiddenDataFinding finding) =>
        finding.Kind == OfficeHiddenDataKind.Provenance ? finding.Provenance?.Carrier : null;

    private static bool DeclaresAi(OfficeHiddenDataFinding finding) => finding.Provenance?.DigitalSourceKind is
        OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;


    private static ToolResultDocument Inspect(ToolSession session, SelectedDocument file) {
        OfficeHiddenDataReport report = OfficeDocument.InspectHiddenData(file.Bytes, file.Name);
        OfficeAccessibilityReport? accessibility = InspectAccessibility(file.Extension, file.Bytes);
        session.LastReview = report;
        ToolArtifact[] artifacts = session.ReplaceArtifacts((JsonReport(file, report, outputName: null), "report"));

        var items = new List<ToolItem>();
        foreach (Category category in Categories) {
            OfficeHiddenDataFinding[] findings = report.Findings.Where(category.Covers).ToArray();
            if (findings.Length == 0) continue;
            string detail = (category.Explanation + " " + Describe(category, findings, IsImage(file.Extension))).Trim();
            items.Add(new ToolItem(category.Id, Title(category, findings.Length), detail, ToolState.Found, category.Group,
                Selectable: category.Selectable, Selected: category.Selectable && category.Preselected));
        }
        if (report.Findings.Any(static finding => finding.Kind == OfficeHiddenDataKind.Provenance && finding.Scope == OfficeHiddenDataScope.Package)) {
            items.Add(new ToolItem("package-credential", "Content Credentials for the whole file",
                "These cover every byte of the file, so any change breaks them. Saving a cleaned copy always removes them, even if you keep the other origin records.",
                ToolState.Warning, ProvenanceText.Group));
        }
        OfficeProvenanceReport provenance = ProvenanceView(report);
        Func<OfficeProvenanceEvidence, string?> assetOf = AssetResolver(report);
        OfficeProvenanceEvidence[] manifests = provenance.Evidence.Where(static evidence => evidence.Manifest != null).Take(3).ToArray();
        for (int index = 0; index < manifests.Length; index++) {
            items.AddRange(ProvenanceText.Explain(manifests[index].Manifest!, index, manifests.Length, manifests[index], assetOf));
        }
        if (accessibility != null) items.AddRange(AccessibilityItems(file.Extension, accessibility));

        var facts = new List<ToolFact> { new("Format", FormatName(file.Extension)) };
        string? author = report.GetFindings(OfficeHiddenDataKind.DocumentProperties)
            .Where(static finding => finding.Description.Contains("Author", StringComparison.OrdinalIgnoreCase) || finding.Description.Contains("modified by", StringComparison.OrdinalIgnoreCase))
            .Concat(report.GetFindings(OfficeHiddenDataKind.PersonalIdentities))
            .Select(static finding => finding.Value).FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value));
        if (author != null) facts.Add(new ToolFact("Name in the file", Clip(author)!, ToolTone.Warn));
        if (report.GetFindings(OfficeHiddenDataKind.ImageMetadata).Any(static finding => finding.Description.Contains("GPS", StringComparison.OrdinalIgnoreCase))) {
            facts.Add(new ToolFact("Photo location", file.Extension is ".pdf" || IsWordProcessingPackage(file.Extension) ? "Stored in a picture" : "Stored in the photo", ToolTone.Warn));
        }
        if (report.GetFindings(OfficeHiddenDataKind.TrackedChanges).Count is int changes and > 0) {
            facts.Add(new ToolFact("Tracked changes", changes.ToString(CultureInfo.InvariantCulture), ToolTone.Warn));
        }
        if (manifests.Length > 0) {
            OfficeC2paManifestSummary first = manifests[0].Manifest!;
            string? madeWith = first.Actions.Select(static action => Clip(action.SoftwareAgent)).FirstOrDefault(static agent => agent != null) ?? Clip(first.ClaimGenerator);
            if (madeWith != null) facts.Add(new ToolFact(manifests.Length > 1 ? "First recorded tool" : "Recorded tool", madeWith));
            if (Clip(first.SignedBy) is { } signer) facts.Add(new ToolFact(manifests.Length > 1 ? "First certificate subject" : "Certificate subject", signer + " · unverified", ToolTone.Warn));
        }
        if (provenance.HasGenerativeAiDeclaration) facts.Add(new ToolFact("Generative AI", "Declared in an origin record", ToolTone.Warn));
        if (accessibility != null) {
            facts.Add(new ToolFact("Accessibility", accessibility.ErrorCount == 0 ? "No errors" : ToolFormat.Count(accessibility.ErrorCount, "error"),
                accessibility.ErrorCount == 0 ? ToolTone.Good : ToolTone.Warn));
        }

        int kinds = items.Count(static item => item.Selectable && item.Id != TableHeaders);
        string origin = provenance.Evidence.Count > 0 ? " " + ProvenanceText.Summary(provenance, ProvenanceText.ManifestSummaries(provenance), assetOf) : string.Empty;
        ToolVerdict verdict = kinds > 0
            ? new ToolVerdict(ToolTone.Warn, "Found " + ToolFormat.Count(kinds, "kind") + " of data that would leave with this file",
                "Ticked items are removed from a copy. Tracked changes, hidden content, attachments and origin records stay unless you tick them, because removing them changes what readers see or what the file tells them." + origin)
            : new ToolVerdict(ToolTone.Good, "Nothing hidden found",
                "This file has none of the properties, comments, hidden content, picture details or origin records that the tool looks for." + origin);
        return new ToolResultDocument(true, verdict, facts.ToArray(), items.ToArray(), artifacts, null, 0);
    }

    /// <summary>Accessibility findings for Word and PowerPoint. Only the safe Word fix is offered as a choice; the rest need a person.</summary>
    private static IEnumerable<ToolItem> AccessibilityItems(string extension, OfficeAccessibilityReport accessibility) {
        if (IsWordProcessingPackage(extension) && accessibility.Findings.Any(static finding => finding.Code == MissingTableHeader)) {
            yield return new ToolItem(TableHeaders, "Mark table header rows",
                "Marks the first row of each data table as a header row, so screen readers announce column names. Tables that look like layout tables are left alone.",
                ToolState.Found, AccessibilityGroup, Selectable: true, Selected: false);
        }
        int shown = 0;
        foreach (IGrouping<string, OfficeAccessibilityFinding> rule in accessibility.Findings.GroupBy(static finding => finding.Code)) {
            if (shown++ == 8) break;
            OfficeAccessibilityFinding first = rule.First();
            int count = rule.Count();
            yield return new ToolItem("a11y-" + rule.Key, AccessibilityTitle(rule.Key) + (count > 1 ? $" ({count})" : string.Empty),
                first.Message + (count > 1 ? " The same applies elsewhere in the file." : string.Empty),
                first.Severity == OfficeAccessibilitySeverity.Error ? ToolState.Warning : ToolState.Info, AccessibilityGroup);
        }
        if (accessibility.Findings.Count == 0) {
            yield return new ToolItem("a11y-ok", "No accessibility problems found", "Titles, language, alternative text, headings, tables, links and contrast passed the checks.", ToolState.Good, AccessibilityGroup);
        }
    }

    private const string MissingTableHeader = "accessibility-missing-table-header";

    private static string AccessibilityTitle(string code) => code switch {
        "accessibility-missing-document-title" => "No document title",
        "accessibility-missing-language" => "No document language",
        "accessibility-missing-alternative-text" => "Pictures without alternative text",
        "accessibility-missing-headings" => "Long document without headings",
        MissingTableHeader => "Tables without a header row",
        "accessibility-layout-table" => "Table used for layout",
        "accessibility-unclear-link-label" => "Unclear link text",
        "accessibility-low-contrast" => "Low text contrast",
        "accessibility-floating-object" => "Floating picture",
        "accessibility-blank-paragraph-spacing" => "Empty paragraphs used for spacing",
        "accessibility-missing-slide-title" => "Slides without a title",
        "accessibility-missing-shape-title" => "Shapes without a title",
        "accessibility-missing-link-tooltip" => "Links without a tooltip",
        "accessibility-chart-color-only-meaning" => "Chart relies on colour alone",
        _ => code
    };

    private static ToolResultDocument Remove(ToolSession session, SelectedDocument file, ToolOptions options) {
        IReadOnlySet<string> selected = options.List("remove");
        if (selected.Count == 0) throw new InvalidOperationException("Select at least one item to remove or fix.");
        var removal = new OfficeHiddenDataOptions();
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (Category category in HiddenData) {
            if (selected.Contains(category.Id)) kinds |= category.Kind;
        }
        removal.Provenance.RemoveC2paManifests = selected.Contains(ProvenanceText.Manifests);
        removal.Provenance.RemoveExternalC2paReferences = selected.Contains(ProvenanceText.References);
        removal.Provenance.RemoveAiSourceMetadata = selected.Contains(ProvenanceText.Declarations);
        if (removal.Provenance.RemoveC2paManifests || removal.Provenance.RemoveExternalC2paReferences || removal.Provenance.RemoveAiSourceMetadata) {
            kinds |= OfficeHiddenDataKind.Provenance;
        }
        removal.Kinds = kinds;
        removal.TrackedChanges = (kinds & OfficeHiddenDataKind.TrackedChanges) != 0 ? OfficeTrackedChangesAction.AcceptAll : OfficeTrackedChangesAction.Keep;
        bool fixHeaders = selected.Contains(TableHeaders) && IsWordProcessingPackage(file.Extension);

        OfficeHiddenDataResult result = OfficeDocument.RemoveHiddenData(file.Bytes, file.Name, removal);
        OfficeHiddenDataReport report = result.Report;
        byte[] output = result.ToArray();
        int headersMarked = 0;
        if (fixHeaders) (output, headersMarked) = MarkTableHeaderRows(output);
        OfficeHiddenDataReport after = OfficeDocument.InspectHiddenData(output, file.Name);

        string outputName = Path.GetFileNameWithoutExtension(file.Name) + "-cleaned" + file.Extension;
        ToolArtifact[] artifacts = session.ReplaceArtifacts(
            (new BrowserConversionArtifact(output, outputName, ContentType(file.Extension)), "primary"),
            (JsonReport(file, report, outputName), "report"));

        var items = new List<ToolItem>();
        int changed = 0, kept = 0;
        foreach (Category category in Categories) {
            OfficeHiddenDataFinding[] findings = report.Findings.Where(category.Covers).ToArray();
            if (findings.Length == 0) continue;
            OfficeHiddenDataFinding[] done = findings.Where(static finding => finding.Outcome is not (OfficeHiddenDataOutcome.Found or OfficeHiddenDataOutcome.Retained)).ToArray();
            OfficeHiddenDataFinding[] retained = findings.Where(static finding => finding.Outcome == OfficeHiddenDataOutcome.Retained).ToArray();
            changed += done.Length;
            kept += retained.Length;
            if (done.Length > 0) {
                items.Add(new ToolItem(category.Id, Title(category, done.Length), Outcomes(done), ToolState.Removed, category.Group));
            }
            if (retained.Length > 0) {
                bool byChoice = category.Selectable && !selected.Contains(category.Id);
                string detail = byChoice
                    ? "Kept because you chose to keep it."
                    : !category.Selectable
                        ? "Kept: these labels don't say the content was made by AI, and only AI labels are removed."
                        : "Kept: " + string.Join(" ", retained.Select(Explain).Where(static note => note.Length > 0).Distinct().Take(3));
                items.Add(new ToolItem(category.Id + "-kept", Title(category, retained.Length), detail.TrimEnd(), byChoice || !category.Selectable ? ToolState.Kept : ToolState.Info, category.Group));
            }
        }
        if (fixHeaders) {
            items.Add(new ToolItem(TableHeaders, "Table header rows", headersMarked > 0
                ? ToolFormat.Count(headersMarked, "table") + " now " + (headersMarked == 1 ? "has" : "have") + " a header row."
                : "No table needed a header row, or every remaining one looks like a layout table.", headersMarked > 0 ? ToolState.Good : ToolState.Info, AccessibilityGroup));
        }
        bool watermark = removal.Provenance.RemoveC2paManifests && report.Findings
            .Where(static finding => finding.Kind == OfficeHiddenDataKind.Provenance && finding.Outcome == OfficeHiddenDataOutcome.Removed)
            .Select(static finding => finding.Provenance?.Manifest).OfType<OfficeC2paManifestSummary>()
            .Any(static manifest => manifest.Actions.Any(static action => ProvenanceText.Watermarked(action.Action)));
        if (watermark) items.Add(new ToolItem("watermark", "Watermark in the content", ProvenanceText.WatermarkNote, ToolState.Warning, ProvenanceText.Group));
        bool image = IsImage(file.Extension);
        items.Add(new ToolItem("content", image ? "Image pixels" : "Visible content",
            (kinds & (OfficeHiddenDataKind.HiddenContent | OfficeHiddenDataKind.TrackedChanges | OfficeHiddenDataKind.Attachments)) == 0
                ? "Not changed. Only the items you ticked were touched."
                : "Changed only where you chose to remove hidden content or attachments, or accept tracked changes.", ToolState.Kept));

        // Names on comments and changes that stay are anonymized, not removed, and retained items stay with a reason;
        // anything else in a selected category that the re-check still finds is reported as not removed.
        Category[] chosen = Categories.Where(category => selected.Contains(category.Id)
            && !report.Findings.Any(finding => category.Covers(finding) && finding.Outcome == OfficeHiddenDataOutcome.Retained)).ToArray();
        int stillThere = after.Findings.Count(finding => finding.Kind != OfficeHiddenDataKind.TrackedChanges
            && chosen.Any(category => category.Covers(finding))
            && !(finding.Kind == OfficeHiddenDataKind.PersonalIdentities && finding.Value == removal.AnonymousAuthor)
            && finding.Reason != OfficeHiddenDataReason.SourceTypeKept);
        string title = changed == 0 && !fixHeaders ? "Nothing was removed" : "Clean copy ready" + (changed > 0 ? " · " + ToolFormat.Count(changed, "item") + " removed or anonymized" : string.Empty);
        string detailText = stillThere > 0
            ? "We checked the copy again and some selected items are still there. See the list below."
            : kept > 0
                ? "We checked the copy again. The items shown as kept remain, each with the reason."
                : "We checked the copy again and found none of the selected items.";
        if (watermark) detailText += " The watermark a removed record mentions is part of the content and is still there.";
        var verdict = new ToolVerdict(changed == 0 && !fixHeaders || stillThere > 0 ? ToolTone.Warn : ToolTone.Good, title, detailText);
        ToolPreview? preview = image ? new ToolPreview("image", 0) : file.Extension == ".pdf" ? new ToolPreview("pdf", 0) : null;
        return new ToolResultDocument(true, verdict,
            [new ToolFact("Removed", changed.ToString(CultureInfo.InvariantCulture), changed > 0 ? ToolTone.Good : null),
             new ToolFact("Kept", kept.ToString(CultureInfo.InvariantCulture), kept > 0 ? ToolTone.Warn : null),
             new ToolFact("Size", ToolFormat.Bytes(output.LongLength))],
            items.ToArray(), artifacts, preview, 0);
    }

    /// <summary>
    /// The page's sentence for a finding that was kept or changed with a reason. Reasons that point at a removal choice name the
    /// checkbox; the others use the engine's own sentence.
    /// </summary>
    private static string Explain(OfficeHiddenDataFinding finding) => finding.Reason switch {
        OfficeHiddenDataReason.ProtectsContentCredentials =>
            "Changing this picture would break its Content Credentials. Tick Content Credentials too to clean both.",
        OfficeHiddenDataReason.SourceTypeKept =>
            "The label that says how the picture was made stays. Tick AI source label to remove it too.",
        OfficeHiddenDataReason.WholeFileCredentialsCannotSurvive =>
            "Content Credentials for the whole file can't stay valid once anything in the file changes, so they go with any clean copy.",
        OfficeHiddenDataReason.KeptByProvenanceChoices =>
            "The record is malformed, or it is a kind you chose to keep.",
        _ => finding.Note ?? string.Empty
    };

    /// <summary>The hidden-data report's provenance findings, viewed as the provenance report the origin explanations read.</summary>
    private static OfficeProvenanceReport ProvenanceView(OfficeHiddenDataReport report) => new(OfficeProvenanceAssetFormat.Unknown,
        report.GetFindings(OfficeHiddenDataKind.Provenance).Select(static finding => finding.Provenance).OfType<OfficeProvenanceEvidence>().ToList());

    /// <summary>Names the picture each record belongs to, from the finding's part, so the explanations can say which one it is.</summary>
    private static Func<OfficeProvenanceEvidence, string?> AssetResolver(OfficeHiddenDataReport report) {
        var assets = new Dictionary<OfficeProvenanceEvidence, string>(ReferenceEqualityComparer.Instance);
        foreach (OfficeHiddenDataFinding finding in report.GetFindings(OfficeHiddenDataKind.Provenance)) {
            if (finding.Provenance is { } evidence && finding.Scope == OfficeHiddenDataScope.EmbeddedImage && finding.Part is { Length: > 0 } part) {
                assets[evidence] = part.TrimStart('/');
            }
        }
        return evidence => assets.TryGetValue(evidence, out string? asset) ? asset : null;
    }

    private static string Place(OfficeHiddenDataFinding finding, bool standaloneImage) => finding.Scope switch {
        OfficeHiddenDataScope.EmbeddedImage when standaloneImage => "the image's metadata",
        OfficeHiddenDataScope.EmbeddedImage when finding.Part is { Length: > 0 } part => "picture " + part.TrimStart('/'),
        OfficeHiddenDataScope.Package => "the whole file",
        _ => finding.Provenance?.Carrier == OfficeProvenanceCarrierKind.C2paManifest ? "an embedded Content Credentials store"
            : finding.Location.Contains("XMP", StringComparison.OrdinalIgnoreCase) ? "the file's XMP metadata" : "the file"
    };

    /// <summary>What was found: a short sample of names or properties, or for origin records, where they sit.</summary>
    private static string Describe(Category category, OfficeHiddenDataFinding[] findings, bool standaloneImage) {
        if (category.Group != ProvenanceText.Group) return Examples(findings);
        string what = category.Id is ProvenanceText.Declarations or SourceLabels
            ? ProvenanceText.SourceTypeText(findings.Select(static finding => finding.Provenance).OfType<OfficeProvenanceEvidence>().ToArray()) + " "
            : category.Id == ProvenanceText.References && Clip(findings[0].Value) is { } uri ? $"({uri}) " : string.Empty;
        string[] places = findings.Select(finding => Place(finding, standaloneImage)).Distinct(StringComparer.Ordinal).ToArray();
        string where = string.Join("; ", places.Take(3)) + (places.Length > 3 ? $"; and {places.Length - 3} more places" : string.Empty);
        string validity = findings.All(static finding => finding.Provenance?.IsStructurallyValid != false) ? string.Empty : " Some look malformed and will be left in place.";
        return $"{what}Found in {where}.{validity}";
    }

    private static string Title(Category category, int count) => count > 1 ? $"{category.Title} ({count})" : category.Title;

    /// <summary>A short, readable sample of what was found, such as the names or properties.</summary>
    private static string Examples(IReadOnlyList<OfficeHiddenDataFinding> findings) {
        string[] samples = findings.Select(static finding => finding.Value is { Length: > 0 } value ? $"{finding.Description}: {Clip(value)}" : finding.Description)
            .Distinct(StringComparer.Ordinal).Take(3).ToArray();
        string more = findings.Count > samples.Length ? $" and {findings.Count - samples.Length} more" : string.Empty;
        string text = "Found: " + string.Join("; ", samples) + more;
        return text.EndsWith('…') || text.EndsWith('.') ? text : text + ".";
    }

    private static string Outcomes(OfficeHiddenDataFinding[] findings) {
        var parts = findings.GroupBy(static finding => finding.Outcome).Select(static group => group.Key switch {
            OfficeHiddenDataOutcome.Anonymized => ToolFormat.Count(group.Count(), "name") + " replaced with an anonymous author",
            OfficeHiddenDataOutcome.Accepted => ToolFormat.Count(group.Count(), "change") + " accepted",
            OfficeHiddenDataOutcome.Rejected => ToolFormat.Count(group.Count(), "change") + " rejected",
            _ => ToolFormat.Count(group.Count(), "item") + " removed"
        });
        string notes = string.Join(" ", findings.Where(static finding => finding.Reason != OfficeHiddenDataReason.None)
            .Select(Explain).Where(static note => note.Length > 0).Distinct().Take(2));
        return (string.Join(", ", parts) + "." + (notes.Length > 0 ? " " + notes : string.Empty)).Trim();
    }

    private static string? Clip(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string text = new string(value.Trim().Where(static ch => !char.IsControl(ch)).ToArray());
        return text.Length <= 60 ? text : text.Substring(0, 59) + "…";
    }

    private static BrowserConversionArtifact JsonReport(SelectedDocument file, OfficeHiddenDataReport report, string? outputName) =>
        new(Encoding.UTF8.GetBytes(OfficeHiddenDataReportSerializer.Serialize(report, file.Name, outputName)),
            Path.GetFileNameWithoutExtension(file.Name) + "-hidden-data-report.json", "application/json");

    /// <summary>The accessibility checker is a Word and PowerPoint feature; other formats have no report.</summary>
    private static OfficeAccessibilityReport? InspectAccessibility(string extension, byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        if (IsWordProcessingPackage(extension)) {
            using WordDocument document = WordDocument.Load(stream, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PackageSecurity = BrowserConversionService.CreateBrowserPackageSecurity() });
            return document.InspectAccessibility();
        }
        if (extension == ".pptx") {
            using PowerPointDocument document = PowerPointDocument.Load(stream, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PackageSecurity = BrowserConversionService.CreateBrowserPackageSecurity() });
            return document.InspectAccessibility();
        }
        return null;
    }

    /// <summary>Word's one automatic accessibility fix, applied to the cleaned copy. Returns the new bytes and how many tables gained a header row.</summary>
    private static (byte[] Output, int Marked) MarkTableHeaderRows(byte[] cleaned) {
        using var stream = new MemoryStream();
        stream.Write(cleaned, 0, cleaned.Length);
        stream.Position = 0;
        using WordDocument document = WordDocument.Load(stream, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadWrite, PackageSecurity = BrowserConversionService.CreateBrowserPackageSecurity() });
        int before = document.InspectAccessibility().Findings.Count(static finding => finding.Code == MissingTableHeader);
        int after = document.FixAccessibility(new WordAccessibilityFixOptions { MarkTableHeaderRows = true }).Findings.Count(static finding => finding.Code == MissingTableHeader);
        using var output = new MemoryStream();
        document.Save(output);
        return (output.ToArray(), Math.Max(0, before - after));
    }

    private static bool IsWordProcessingPackage(string extension) => extension == ".docx";

    private static bool IsImage(string extension) => extension is ".png" or ".jpg" or ".jpeg" or ".webp";

    private static string FormatName(string extension) => extension switch {
        ".docx" => "Word document",
        ".xlsx" => "Excel workbook",
        ".pptx" => "PowerPoint presentation",
        ".pdf" => "PDF",
        ".jpg" or ".jpeg" => "JPEG image",
        ".png" => "PNG image",
        ".webp" => "WebP image",
        _ => extension.TrimStart('.').ToUpperInvariant()
    };

    private static string ContentType(string extension) => extension switch {
        ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pdf" => "application/pdf",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".png" => "image/png",
        ".webp" => "image/webp",
        _ => "application/octet-stream"
    };
}
