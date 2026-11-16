using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAccessibilityInspector {
    private static readonly HashSet<string> KnownAnnotationSubtypes = new(StringComparer.Ordinal) {
        "Text", "Link", "FreeText", "Line", "Square", "Circle", "Polygon", "PolyLine", "Highlight", "Underline", "Squiggly",
        "StrikeOut", "Stamp", "Caret", "Ink", "Popup", "FileAttachment", "Sound", "Movie", "Widget", "Screen", "PrinterMark",
        "TrapNet", "Watermark", "3D", "Redact", "RichMedia", "Projection"
    };

    // Every condition whose evidence comes from the page content scan, including the font checks it feeds.
    private static readonly string[] ContentScanConditions = new[] { "01-003", "01-004", "01-005", "11-001", "11-002", "30-001", "30-002" }
        .Concat(PdfAccessibilityConditions.All
            .Where(static condition => condition.Checkpoint == "31" && condition.Coverage == PdfAccessibilityCoverage.Automatic)
            .Select(static condition => condition.ConditionId))
        .ToArray();

    private void CheckPages(StructureModel structure) {
        var fonts = new Dictionary<PdfDictionary, (bool Rendered, int? ObjectNumber, int Page)>(PdfAccessibilityPageContent.ReferenceEqualityComparer<PdfDictionary>.Instance);
        var formInvocations = new Dictionary<int, int>();
        var formsWithMcids = new HashSet<int>();
        var formPage = new Dictionary<int, int>();
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        for (int index = 0; index < pages.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfReadPage page = pages[index];
            int pageNumber = index + 1;
            CheckAnnotations(page, pageNumber, structure);
            PdfAccessibilityPageContent content;
            try {
                content = page.ScanAccessibilityContent((stream, mcid) => structure.IsLinked(page.ObjectNumber, stream, mcid), _cancellationToken);
            } catch (Exception exception) when (IsReadFailure(exception)) {
                foreach (string id in ContentScanConditions) {
                    Unknown(id, null, "The page content could not be inspected completely: " + exception.Message, pageNumber);
                }
                continue;
            }
            Applies("01-003");
            Applies("01-004");
            Applies("01-005");
            if (content.ArtifactsInTaggedContent > 0) {
                Fail("01-003", "7.1-1", Count(content.ArtifactsInTaggedContent, "Artifact sequence") + " inside tagged content.", pageNumber, page.ObjectNumber);
            }
            if (content.TaggedContentInArtifacts > 0) {
                Fail("01-004", "7.1-2", Count(content.TaggedContentInArtifacts, "tagged sequence") + " inside Artifact content.", pageNumber, page.ObjectNumber);
            }
            if (content.UntaggedItems > 0) {
                Fail("01-005", "7.1-3", Count(content.UntaggedItems, "content item") + " neither tagged nor marked as Artifact. " +
                    "Tag real content in the structure tree and mark decoration as Artifact.", pageNumber, page.ObjectNumber);
            }
            foreach (string language in content.PropertyLanguages) CheckLanguageValue(language, pageNumber, null, "a marked-content property list", page.ObjectNumber);
            CheckContentLanguage(content, page, pageNumber, structure);
            foreach (KeyValuePair<PdfDictionary, bool> font in content.Fonts) {
                content.FontObjectNumbers.TryGetValue(font.Key, out int? fontObject);
                fonts[font.Key] = fonts.TryGetValue(font.Key, out var previous)
                    ? (previous.Rendered || font.Value, previous.ObjectNumber ?? fontObject, previous.Page)
                    : (font.Value, fontObject, pageNumber);
            }
            foreach (KeyValuePair<int, int> invocation in content.FormInvocations) {
                formInvocations[invocation.Key] = formInvocations.TryGetValue(invocation.Key, out int count) ? count + invocation.Value : invocation.Value;
                if (!formPage.ContainsKey(invocation.Key)) formPage[invocation.Key] = pageNumber;
            }
            formsWithMcids.UnionWith(content.FormsWithMcids);
            if (content.FormInvocations.Count > 0) Applies("30-001");
            foreach (int reference in content.ReferenceXObjects) {
                Fail("30-001", "7.20-1", "A reference XObject is painted; embed the referenced content instead.", pageNumber, reference);
            }
        }
        foreach (int form in formsWithMcids) {
            Applies("30-002");
            if (formInvocations.TryGetValue(form, out int count) && count > 1) {
                Fail("30-002", "7.20-2", "A Form XObject containing MCIDs is painted " + count.ToString(CultureInfo.InvariantCulture) +
                    " times; each painted copy needs its own structure, so give each use its own XObject.", formPage[form], form);
            }
        }
        CheckFonts(fonts);
    }

    private static string Count(int count, string noun) =>
        count.ToString(CultureInfo.InvariantCulture) + " " + noun + (count == 1 ? " is" : "s are");

    private void CheckContentLanguage(PdfAccessibilityPageContent content, PdfReadPage page, int pageNumber, StructureModel structure) {
        if (content.TextMcids.Count > 0 || content.UntaggedTextWithoutLanguage > 0) Applies("11-001");
        if (!_catalogHasLanguage && content.UntaggedTextWithoutLanguage > 0) {
            Fail("11-001", "7.2-34", Count(content.UntaggedTextWithoutLanguage, "untagged text item") + " without a determinable language. Set Lang on the catalog.",
                pageNumber, page.ObjectNumber);
        }
        int missing = 0;
        foreach ((int? stream, int mcid) in content.TextMcids) {
            if (_catalogHasLanguage) continue;
            if (!HasLanguage(structure.Owner(page.ObjectNumber, stream, mcid))) missing++;
        }
        if (missing > 0) {
            Fail("11-001", "7.2-34", Count(missing, "tagged text sequence") + " without a determinable language. Set Lang on the catalog or on the structure elements.",
                pageNumber, page.ObjectNumber);
        }
        foreach ((int? stream, int? mcid, string key) in content.PropertyTextWithoutLang) {
            Applies("11-002");
            StructNode? owner = mcid is int id ? structure.Owner(page.ObjectNumber, stream, id) : null;
            if (!HasLanguage(owner)) {
                Fail("11-002", key == "ActualText" ? "7.2-30" : key == "Alt" ? "7.2-31" : "7.2-32",
                    "Marked content carries " + key + " text without a determinable language.", pageNumber, page.ObjectNumber, owner);
            }
        }
    }

    private void CheckAnnotations(PdfReadPage page, int pageNumber, StructureModel structure) {
        PdfDictionary pageDictionary = page.PageDictionary;
        PdfArray? annotations = ResolveArray(pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotsObject) ? annotsObject : null);
        if (annotations is null || annotations.Items.Count == 0) return;
        Applies("28-008");
        string? tabs = ReadName(pageDictionary, "Tabs");
        if (tabs is null) {
            Fail("28-008", "7.18.3-1", "The page has annotations but no Tabs entry; set Tabs to S.", pageNumber, page.ObjectNumber);
        } else {
            Applies("28-009");
            if (tabs != "S") Fail("28-009", "7.18.3-1", "The page has annotations and Tabs is " + tabs + "; set it to S.", pageNumber, page.ObjectNumber);
        }
        PdfPageGeometry geometry = page.GetGeometry();
        PdfPageBox? crop = geometry.CropBox ?? geometry.MediaBox;
        foreach (PdfObject item in annotations.Items) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (ResolveDict(item) is not PdfDictionary annotation) continue;
            int? objectNumber = ObjectNumberOf(item);
            string subtype = ReadName(annotation, "Subtype") ?? string.Empty;
            bool hidden = ((ReadInteger(annotation, "F") ?? 0) & 2) == 2;
            bool outside = IsOutside(annotation, crop);
            StructNode? owner = FindAnnotationOwner(objectNumber, annotation, structure);
            string? alt = owner is null ? null : ReadText(owner.Dictionary, "Alt");
            string? contents = ReadText(annotation, "Contents");

            if (subtype == "TrapNet") {
                Applies("28-007");
                Fail("28-007", "7.18.2-1", "TrapNet annotations are not allowed.", pageNumber, objectNumber);
            }
            if (contents is not null) {
                Applies("11-004");
                if (!_catalogHasLanguage && string.IsNullOrEmpty(owner?.Language)) {
                    Fail("11-004", "7.2-24", "The annotation Contents has no determinable language.", pageNumber, objectNumber, owner);
                }
            }
            if (subtype == "PrinterMark") {
                Applies("28-017");
                if (owner is not null && !hidden && !outside) {
                    Fail("28-017", "7.18.8-1", "A PrinterMark annotation is part of the logical structure; it must be an artifact.", pageNumber, objectNumber, owner);
                }
            }
            if (subtype == "FileAttachment" && annotation.Items.TryGetValue("FS", out PdfObject? fileSpecification)) {
                CheckFileSpecification(fileSpecification, "28-016", pageNumber);
            }
            CheckMediaClips(annotation, pageNumber, objectNumber);
            if (hidden || outside) continue;

            string nestingId = KnownAnnotationSubtypes.Contains(subtype) ? "28-002" : "28-006";
            if (subtype is not ("Widget" or "Link" or "PrinterMark")) {
                Applies(nestingId);
                if (owner?.StandardType != "Annot") {
                    Fail(nestingId, "7.18.1-1", "The " + (subtype.Length == 0 ? "annotation" : subtype + " annotation") + " is not inside an Annot structure element.", pageNumber, objectNumber, owner);
                }
            }
            if (subtype != "Widget") {
                string descriptionId = KnownAnnotationSubtypes.Contains(subtype) ? "28-004" : "28-006";
                Applies(descriptionId);
                if (string.IsNullOrEmpty(contents) && string.IsNullOrEmpty(alt)) {
                    Fail(descriptionId, "7.18.1-2", "The " + (subtype.Length == 0 ? "annotation" : subtype + " annotation") + " has no Contents and its structure element has no Alt.", pageNumber, objectNumber, owner);
                }
            }
            if (subtype == "Widget") {
                Applies("28-010");
                if (owner?.StandardType != "Form") {
                    Fail("28-010", "7.18.4-1", "The widget annotation is not inside a Form structure element.", pageNumber, objectNumber, owner);
                }
                PdfDictionary field = ResolveField(annotation);
                if (ReadName(field, "FT") is not null || annotation.Items.ContainsKey("Parent") || annotation.Items.ContainsKey("T")) {
                    Applies("28-005");
                    string? tooltip = ReadText(field, "TU");
                    if (string.IsNullOrEmpty(tooltip) && string.IsNullOrEmpty(alt)) {
                        Fail("28-005", "7.18.1-3", "The form field has no TU (tooltip) and its Form element has no Alt; screen readers have no name for it.", pageNumber, objectNumber, owner);
                    }
                    if (tooltip is not null) {
                        Applies("11-005");
                        if (!_catalogHasLanguage && string.IsNullOrEmpty(owner?.Language)) {
                            Fail("11-005", "7.2-25", "The form field TU text has no determinable language; set Lang on the catalog or the Form element.", pageNumber, objectNumber, owner);
                        }
                    }
                }
            }
            if (subtype == "Link") {
                Applies("28-011");
                Applies("28-012");
                if (owner?.StandardType != "Link") {
                    Fail("28-011", "7.18.5-1", "The link annotation is not inside a Link structure element.", pageNumber, objectNumber, owner);
                }
                if (string.IsNullOrEmpty(contents)) {
                    Fail("28-012", "7.18.5-2", "The link annotation has no Contents describing where it goes.", pageNumber, objectNumber, owner);
                }
            }
        }
    }

    private bool HasAlternateDescription(PdfDictionary clip) {
        // Alt is a multi-language text array of language and text pairs.
        if (ResolveArray(clip.Items.TryGetValue("Alt", out PdfObject? altObject) ? altObject : null) is not PdfArray alt) return false;
        for (int i = 1; i < alt.Items.Count; i += 2) {
            if (Resolve(alt.Items[i]) is PdfStringObj text && !string.IsNullOrWhiteSpace(text.Value)) return true;
        }
        return false;
    }

    private StructNode? FindAnnotationOwner(int? objectNumber, PdfDictionary annotation, StructureModel structure) {
        if (objectNumber is int number && structure.AnnotationOwners.TryGetValue(number, out StructNode? owner)) return owner;
        if (ReadInteger(annotation, "StructParent") is int key && structure.ParentTree.TryGetValue(key, out PdfObject? value) &&
            ResolveDict(value) is PdfDictionary element) {
            return structure.NodesByDictionary.TryGetValue(element, out StructNode? node) ? node : null;
        }
        return null;
    }

    private PdfDictionary ResolveField(PdfDictionary widget) {
        // A widget merged with its field carries T or FT itself; otherwise its Parent is the terminal field.
        if (widget.Items.ContainsKey("T") || widget.Items.ContainsKey("FT")) return widget;
        return ResolveDict(widget.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null) ?? widget;
    }

    private bool IsOutside(PdfDictionary annotation, PdfPageBox? crop) {
        if (crop is null || ResolveArray(annotation.Items.TryGetValue("Rect", out PdfObject? rectObject) ? rectObject : null) is not PdfArray rect || rect.Items.Count < 4) return false;
        double[] values = rect.Items.Take(4).Select(item => Resolve(item) is PdfNumber number ? number.Value : double.NaN).ToArray();
        if (values.Any(double.IsNaN)) return false;
        double left = Math.Min(values[0], values[2]), right = Math.Max(values[0], values[2]);
        double bottom = Math.Min(values[1], values[3]), top = Math.Max(values[1], values[3]);
        return right <= crop.Left || left >= crop.Right || top <= crop.Bottom || bottom >= crop.Top;
    }

    private void CheckMediaClips(PdfDictionary annotation, int pageNumber, int? annotationObject) {
        var actions = new List<PdfObject>();
        if (annotation.Items.TryGetValue("A", out PdfObject? action)) actions.Add(action);
        if (ResolveDict(annotation.Items.TryGetValue("AA", out PdfObject? additional) ? additional : null) is PdfDictionary triggers) actions.AddRange(triggers.Items.Values);
        var visited = new HashSet<PdfDictionary>(PdfAccessibilityPageContent.ReferenceEqualityComparer<PdfDictionary>.Instance);
        foreach (PdfObject item in actions) VisitAction(item, 0);

        void VisitAction(PdfObject? value, int depth) {
            if (depth > 32 || ResolveDict(value) is not PdfDictionary dictionary || !visited.Add(dictionary)) return;
            if (ReadName(dictionary, "S") == "Rendition") VisitRendition(dictionary.Items.TryGetValue("R", out PdfObject? rendition) ? rendition : null, depth + 1);
            PdfObject? next = dictionary.Items.TryGetValue("Next", out PdfObject? nextObject) ? Resolve(nextObject) : null;
            if (next is PdfArray nextActions) foreach (PdfObject nextAction in nextActions.Items) VisitAction(nextAction, depth + 1);
            else VisitAction(next, depth + 1);
        }

        void VisitRendition(PdfObject? value, int depth) {
            if (depth > 32 || ResolveDict(value) is not PdfDictionary rendition || !visited.Add(rendition)) return;
            if (ReadName(rendition, "S") == "SR") {
                if (ResolveArray(rendition.Items.TryGetValue("R", out PdfObject? choices) ? choices : null) is PdfArray list) {
                    foreach (PdfObject choice in list.Items) VisitRendition(choice, depth + 1);
                }
                return;
            }
            if (ResolveDict(rendition.Items.TryGetValue("C", out PdfObject? clipObject) ? clipObject : null) is not PdfDictionary clip ||
                ReadName(clip, "S") != "MCD") return;
            Applies("28-014");
            Applies("28-015");
            if (ReadText(clip, "CT") is null && ReadName(clip, "CT") is null) {
                Fail("28-014", "7.18.6.2-1", "The media clip data dictionary has no CT (content type) entry.", pageNumber, ObjectNumberOf(clipObject) ?? annotationObject);
            }
            if (!HasAlternateDescription(clip)) {
                Fail("28-015", "7.18.6.2-2", "The media clip data dictionary has no Alt entry describing the media.", pageNumber, ObjectNumberOf(clipObject) ?? annotationObject);
            }
        }
    }
}
