using OfficeIMO.Drawing;
using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Word.Rtf;

/// <content>
/// Provides result-bearing Word/RTF conversion APIs and fidelity analysis.
/// </content>
public static partial class WordRtfConverterExtensions {
    internal static RtfConversionResult<RtfDocument> ConvertToRtfWithEvidence(this WordDocument document) {
        ArgumentNullException.ThrowIfNull(document);
        RtfDocument converted = document.ConvertToRtfDocument();
        var report = new RtfConversionReport();
        AddWordToRtfDiagnostics(document, report);
        return new RtfConversionResult<RtfDocument>(converted, report);
    }

    internal static RtfConversionResult<WordDocument> ConvertToWordWithEvidence(this RtfDocument document) {
        ArgumentNullException.ThrowIfNull(document);
        var report = new RtfConversionReport();
        WordDocument converted = ToWordDocumentCore(document, report);
        AddRtfToWordDiagnostics(document, report);
        return new RtfConversionResult<WordDocument>(converted, report);
    }

    internal static RtfConversionResult<WordDocument> ConvertToWordWithEvidence(
        this RtfReadResult readResult,
        string? sourcePath = null) {
        ArgumentNullException.ThrowIfNull(readResult);
        RtfConversionResult<WordDocument> converted = readResult.Document.ConvertToWordWithEvidence();
        var report = new RtfConversionReport();
        report.AddReadDiagnostics(readResult.Diagnostics, sourcePath);
        report.Merge(converted.Report);
        return new RtfConversionResult<WordDocument>(converted.Value, report);
    }

    private static void AddRtfToWordDiagnostics(RtfDocument document, RtfConversionReport report) {
        report.Merge(document.GetStyleConversionDiagnostics());
        AddRtfParagraphFrameDiagnostics(document, report);
        AddRtfTableCellGapDiagnostics(document, report);
        AddRtfNotePolicyDiagnostics(document, report);
        if (document.Styles.Count > 0) {
            report.Add(
                RtfConversionSeverity.Information,
                "RtfWordStylesMapped",
                "RTF paragraph, character, and table stylesheet definitions were mapped to Word styles.",
                RtfConversionAction.Preserved,
                feature: "stylesheet",
                count: document.Styles.Count);
        }

        RtfDocumentWriter.EffectiveListTables effectiveLists = RtfDocumentWriter.BuildEffectiveListTables(document);
        int listStructureCount = effectiveLists.Definitions.Count + effectiveLists.Overrides.Count;
        if (listStructureCount > 0) {
            report.Add(
                RtfConversionSeverity.Information,
                "RtfWordListDefinitionsMapped",
                "RTF list definitions, overrides, levels, and paragraph bindings were mapped to Word numbering.",
                RtfConversionAction.Preserved,
                feature: "listtable",
                count: listStructureCount);
        }

        int objectCount = 0;
        int shapeCount = 0;
        int omittedImageCount = 0;
        int normalizedImageCount = 0;
        var visitedNotes = new HashSet<RtfNote>();
        IEnumerable<IRtfBlock> blocks = document.Sections.Count > 0
            ? document.Sections.SelectMany(section => section.Blocks)
            : document.Blocks;
        foreach (IRtfBlock block in blocks) {
            CountUnsupportedRtfBlock(block, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
        }
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                CountUnsupportedRtfBlock(paragraph, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
            }
        }
        foreach (RtfNote note in document.Notes) {
            CountUnsupportedRtfNote(note, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
        }

        AddUnsupportedRtfDiagnostics(report, objectCount, shapeCount, omittedImageCount, normalizedImageCount);
    }

    private static void CountUnsupportedRtfBlock(
        IRtfBlock block,
        HashSet<RtfNote> visitedNotes,
        ref int objectCount,
        ref int shapeCount,
        ref int omittedImageCount,
        ref int normalizedImageCount) {
        switch (block) {
            case RtfObject:
                objectCount++;
                break;
            case RtfShape shape:
                shapeCount++;
                foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) {
                    CountUnsupportedRtfInlines(paragraph.Inlines, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                }
                break;
            case RtfParagraph paragraph:
                CountUnsupportedRtfInlines(paragraph.Inlines, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                break;
            case RtfImage image:
                CountRtfImage(image, ref omittedImageCount, ref normalizedImageCount);
                break;
            case RtfTable table:
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        foreach (IRtfBlock child in cell.Blocks) {
                            CountUnsupportedRtfBlock(child, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                        }
                    }
                }
                break;
        }
    }

    private static void CountUnsupportedRtfInlines(
        IReadOnlyList<IRtfInline> inlines,
        HashSet<RtfNote> visitedNotes,
        ref int objectCount,
        ref int shapeCount,
        ref int omittedImageCount,
        ref int normalizedImageCount) {
        foreach (IRtfInline inline in inlines) {
            switch (inline) {
                case RtfObject:
                    objectCount++;
                    break;
                case RtfShape:
                    shapeCount++;
                    break;
                case RtfImage image:
                    CountRtfImage(image, ref omittedImageCount, ref normalizedImageCount);
                    break;
                case RtfField field:
                    CountUnsupportedRtfInlines(field.Result.Inlines, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                    break;
                case RtfRun run when run.Note != null:
                    CountUnsupportedRtfNote(run.Note, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                    break;
                case RtfGeneratedText generatedText when generatedText.Note != null:
                    CountUnsupportedRtfNote(generatedText.Note, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
                    break;
            }
        }
    }

    private static void CountUnsupportedRtfNote(
        RtfNote note,
        HashSet<RtfNote> visitedNotes,
        ref int objectCount,
        ref int shapeCount,
        ref int omittedImageCount,
        ref int normalizedImageCount) {
        if (!visitedNotes.Add(note)) return;
        foreach (RtfParagraph paragraph in note.Paragraphs) {
            CountUnsupportedRtfBlock(paragraph, visitedNotes, ref objectCount, ref shapeCount, ref omittedImageCount, ref normalizedImageCount);
        }
    }

    private static void CountRtfImage(RtfImage image, ref int omittedImageCount, ref int normalizedImageCount) {
        if (!CanWriteToWord(image)) {
            omittedImageCount++;
        } else if (image.Format == RtfImageFormat.Dib) {
            normalizedImageCount++;
        }
    }

    private static void AddUnsupportedRtfDiagnostics(
        RtfConversionReport report,
        int objectCount,
        int shapeCount,
        int omittedImageCount,
        int normalizedImageCount) {
        if (objectCount > 0) {
            report.Add(
                RtfConversionSeverity.Warning,
                "RtfWordObjectsOmitted",
                "RTF embedded and linked objects are not represented by the Word bridge.",
                RtfConversionAction.Omitted,
                feature: "object",
                count: objectCount);
        }

        if (shapeCount > 0) {
            report.Add(
                RtfConversionSeverity.Warning,
                "RtfWordShapesOmitted",
                "RTF drawing shapes are not represented by the Word bridge.",
                RtfConversionAction.Omitted,
                feature: "shp",
                count: shapeCount);
        }

        if (normalizedImageCount > 0) {
            report.Add(
                RtfConversionSeverity.Information,
                "RtfWordDibImagesNormalized",
                "RTF device-independent bitmap images were normalized to PNG for Word compatibility.",
                RtfConversionAction.Substituted,
                feature: "dibitmap",
                count: normalizedImageCount);
        }

        if (omittedImageCount > 0) {
            report.Add(
                RtfConversionSeverity.Warning,
                "RtfWordImagesOmitted",
                "RTF images with unsupported formats, invalid payloads or invalid visible dimensions are not represented by the Word bridge.",
                RtfConversionAction.Omitted,
                feature: "pict",
                count: omittedImageCount);
        }
    }

    private static void AddWordToRtfDiagnostics(WordDocument document, RtfConversionReport report) {
        List<WordStoryRootCandidate> storyRoots = EnumerateConvertibleWordStoryRoots(document).ToList();
        AddWordParagraphLayoutDiagnostics(document, storyRoots, report);
        var footnoteIds = new HashSet<long>(storyRoots
            .Select(candidate => candidate.Root)
            .Where(note => note.IsW("footnote") && WordXml.ParseInt64(note.WAttr("id")) != null)
            .Select(note => WordXml.ParseInt64(note.WAttr("id"))!.Value));
        var endnoteIds = new HashSet<long>(storyRoots
            .Select(candidate => candidate.Root)
            .Where(note => note.IsW("endnote") && WordXml.ParseInt64(note.WAttr("id")) != null)
            .Select(note => WordXml.ParseInt64(note.WAttr("id"))!.Value));
        var commentIds = new HashSet<string>(storyRoots
            .Select(candidate => candidate.Root)
            .Where(comment => comment.IsW("comment") && !string.IsNullOrWhiteSpace(comment.WAttr("id")))
            .Select(comment => comment.WAttr("id")!), StringComparer.Ordinal);
        List<WordElement> elements = EnumerateWordElements(document.Elements)
            .Concat(EnumerateHeaderFooterElements(document))
            .Concat(EnumerateNoteElements(document, footnoteIds, endnoteIds))
            .Concat(EnumerateCommentElements(document, commentIds))
            .ToList();
        AddWordTableCellSpacingDiagnostics(elements, report);
        int equationCount = elements
            .Count(element => element is WordEquation);
        if (equationCount > 0) {
            report.Add(
                RtfConversionSeverity.Information,
                "WordRtfEquationsMappedToEqFields",
                "Word equations were mapped to native RTF EQ fields with cached visible text.",
                RtfConversionAction.Substituted,
                feature: "equation",
                count: equationCount);
        }

        var unsupported = elements
            .Where(IsUnsupportedWordElement)
            .GroupBy(element => element.GetType().Name, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        foreach (IGrouping<string, WordElement> group in unsupported) {
            report.Add(
                RtfConversionSeverity.Warning,
                "WordRtfElementOmitted",
                "Word element is not represented by the RTF bridge.",
                RtfConversionAction.Omitted,
                feature: group.Key,
                count: group.Count());
        }

        int omittedImageCount = 0;
        int normalizedImageCount = 0;
        int flattenedAnimationCount = 0;
        foreach ((WordImage? Image, bool OmittedByConverter) candidate in EnumerateWordImageCandidates(document, storyRoots)) {
            if (candidate.OmittedByConverter || candidate.Image == null) {
                omittedImageCount++;
                continue;
            }
            WordImage image = candidate.Image;
            RtfImage? converted = CreateRtfImage(
                image,
                out OfficeImageFormat sourceFormat,
                out bool animationDiscarded);
            if (converted == null) {
                omittedImageCount++;
            } else if (converted.Format == RtfImageFormat.Png &&
                       sourceFormat != OfficeImageFormat.Png) {
                normalizedImageCount++;
            }
            if (animationDiscarded) flattenedAnimationCount++;
        }

        if (normalizedImageCount > 0) {
            report.Add(
                RtfConversionSeverity.Information,
                "WordRtfImagesNormalized",
                "Word raster images that RTF cannot embed directly were normalized to PNG.",
                RtfConversionAction.Substituted,
                feature: "image",
                count: normalizedImageCount);
        }
        if (flattenedAnimationCount > 0) {
            report.Add(
                RtfConversionSeverity.Warning,
                "WordRtfImageAnimationFlattened",
                "Animated Word images were flattened to a selected static frame during PNG normalization.",
                RtfConversionAction.Flattened,
                feature: "image",
                count: flattenedAnimationCount);
        }
        if (omittedImageCount > 0) {
            report.Add(
                RtfConversionSeverity.Warning,
                "WordRtfImagesOmitted",
                "Word images with unsupported, external, unavailable, or invalid payloads are not represented by the RTF bridge.",
                RtfConversionAction.Omitted,
                feature: "image",
                count: omittedImageCount);
        }
    }

    private static IEnumerable<WordElement> EnumerateWordElements(IEnumerable<WordElement> elements) {
        foreach (WordElement element in elements) {
            yield return element;
            if (element is WordParagraph paragraph) {
                if (paragraph.IsShape) yield return paragraph.Shape!;
                if (paragraph.IsChart) yield return paragraph.Chart!;
                if (paragraph.IsSmartArt) yield return paragraph.SmartArt!;
                if (paragraph.IsTextBox) yield return paragraph.TextBox!;
                if (paragraph.IsEquation) yield return paragraph.Equation!;
                if (paragraph.IsStructuredDocumentTag) yield return paragraph.StructuredDocumentTag!;
            }
            if (!(element is WordTable table)) continue;
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.GetCells(readOnly: true)) {
                    foreach (WordElement child in EnumerateWordElements(cell.Elements)) {
                        yield return child;
                    }
                }
            }
        }
    }

    private static IEnumerable<(WordImage? Image, bool OmittedByConverter)> EnumerateWordImageCandidates(
        WordDocument document,
        IEnumerable<WordStoryRootCandidate> storyRoots) {
        var visitedParagraphs = new HashSet<OfficeXmlNode>();
        var visitedRuns = new HashSet<OfficeXmlNode>();
        foreach (WordStoryRootCandidate storyRoot in storyRoots) {
            foreach (OfficeXmlNode paragraph in storyRoot.Root.ModelDescendants("p")) {
                if (!visitedParagraphs.Add(paragraph)) continue;
                bool paragraphOmitted = storyRoot.OmittedByConverter ||
                                        !IsParagraphConvertedFromStoryRoot(storyRoot.Root, paragraph);
                foreach ((OfficeXmlNode Run, bool OmittedByConverter) runCandidate in EnumerateConvertibleWordRunsWithFieldState(paragraph)) {
                    if (!visitedRuns.Add(runCandidate.Run)) continue;
                    WordParagraph candidate = document.RunView(paragraph, runCandidate.Run);
                    if (!TryEnumerateImages(candidate, out IReadOnlyList<WordImage> images)) {
                        int unavailableImages = Math.Max(1, CountImageElements(runCandidate.Run));
                        for (int index = 0; index < unavailableImages; index++) {
                            yield return (null, true);
                        }
                        continue;
                    }
                    foreach (WordImage image in images) {
                        yield return (image, paragraphOmitted || runCandidate.OmittedByConverter);
                    }
                }
            }
        }
    }

    private static int CountImageElements(OfficeXmlNode run) {
        int count = 0;
        foreach (OfficeXmlNode element in EnumerateEffectiveRunContent(run)) {
            IEnumerable<OfficeXmlNode> drawings = element.IsW("drawing")
                ? new[] { element }
                : element.ModelDescendants("drawing");
            count += drawings.Count(drawing => drawing.DescendantElements(WordXml.Pic, "pic").Any(WordConverterNodes.IsModelElement));

            IEnumerable<OfficeXmlNode> shapes = element.Is(WordXml.V, "shape")
                ? new[] { element }
                : element.DescendantElements(WordXml.V, "shape").Where(WordConverterNodes.IsModelElement);
            count += shapes.Count(shape => shape.Child(WordXml.V, "imagedata").Exists);
        }
        return count;
    }

    private static IEnumerable<OfficeXmlNode> EnumerateEffectiveRunContent(OfficeXmlNode run) {
        foreach (OfficeXmlNode child in run.Elements()) {
            if (!child.Is(WordXml.Mc, "AlternateContent")) {
                yield return child;
                continue;
            }

            OfficeXmlNode branch = WordConverterNodes.SelectAlternateContentBranch(child);
            if (branch.IsNone) continue;
            foreach (OfficeXmlNode branchChild in branch.Elements()) {
                yield return branchChild;
            }
        }
    }

    private static bool TryEnumerateImages(WordParagraph paragraph, out IReadOnlyList<WordImage> images) {
        try {
            images = paragraph.EnumerateImages().ToList();
            return true;
        } catch (ArgumentOutOfRangeException) {
            // A dangling image relationship can exist in malformed or revision-deleted content.
            // Diagnostics must not make an otherwise successful conversion fail.
            images = Array.Empty<WordImage>();
            return false;
        }
    }

    private static bool IsParagraphConvertedFromStoryRoot(OfficeXmlNode root, OfficeXmlNode paragraph) {
        if (root.IsW("document")) return true;
        // The Word-to-RTF bridge reads only the direct paragraph collections exposed by
        // headers, footers, notes, and comments. Nested table/text-box paragraphs are inventoried
        // for diagnostics but are not emitted by those specialized story converters.
        return paragraph.Parent == root;
    }

    private static IEnumerable<WordStoryRootCandidate> EnumerateConvertibleWordStoryRoots(WordDocument document) {
        OfficeXmlNode documentRoot = document.DocumentNode();
        if (documentRoot.IsNone) yield break;

        var rootOmissions = new Dictionary<OfficeXmlNode, bool>();
        rootOmissions.Add(documentRoot, false);
        var convertedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        if (document.Sections.Count > 0) {
            OfficeXmlNode sectionProperties = document.Sections[0].Node();
            foreach (OfficeXmlNode reference in sectionProperties.WElements("headerReference")) {
                string? id = reference.Attribute(WordXml.R, "id");
                if (!string.IsNullOrWhiteSpace(id)) convertedRelationshipIds.Add(id!);
            }
            foreach (OfficeXmlNode reference in sectionProperties.WElements("footerReference")) {
                string? id = reference.Attribute(WordXml.R, "id");
                if (!string.IsNullOrWhiteSpace(id)) convertedRelationshipIds.Add(id!);
            }
        }
        foreach ((OfficeXmlNode root, string relationshipId) in document.HeaderFooterRootsWithRelationshipIds()) {
            rootOmissions[root] = !convertedRelationshipIds.Contains(relationshipId);
        }

        bool added;
        do {
            added = false;
            var referencedCommentIds = new Dictionary<string, bool>(StringComparer.Ordinal);
            var referencedFootnoteIds = new Dictionary<long, bool>();
            var referencedEndnoteIds = new Dictionary<long, bool>();
            foreach (KeyValuePair<OfficeXmlNode, bool> candidate in rootOmissions.ToList()) {
                foreach (string id in CollectReferencedCommentIds(new[] { candidate.Key })) {
                    RecordStoryReference(referencedCommentIds, id, candidate.Value);
                }
                foreach (OfficeXmlNode reference in candidate.Key.ModelDescendants("footnoteReference")) {
                    if (WordXml.ParseInt64(reference.WAttr("id")) is long id) RecordStoryReference(referencedFootnoteIds, id, candidate.Value);
                }
                foreach (OfficeXmlNode reference in candidate.Key.ModelDescendants("endnoteReference")) {
                    if (WordXml.ParseInt64(reference.WAttr("id")) is long id) RecordStoryReference(referencedEndnoteIds, id, candidate.Value);
                }
            }

            foreach (OfficeXmlNode comment in document.CommentsNode().WElements("comment")) {
                if (comment.WAttr("id") is string id && referencedCommentIds.TryGetValue(id, out bool omitted)) {
                    added |= RecordStoryRoot(rootOmissions, comment, omitted);
                }
            }
            foreach (OfficeXmlNode note in document.FootnotesNode().WElements("footnote")) {
                if (WordXml.ParseInt64(note.WAttr("id")) is long id && referencedFootnoteIds.TryGetValue(id, out bool omitted)) {
                    added |= RecordStoryRoot(rootOmissions, note, omitted);
                }
            }
            foreach (OfficeXmlNode note in document.EndnotesNode().WElements("endnote")) {
                if (WordXml.ParseInt64(note.WAttr("id")) is long id && referencedEndnoteIds.TryGetValue(id, out bool omitted)) {
                    added |= RecordStoryRoot(rootOmissions, note, omitted);
                }
            }
        } while (added);

        foreach (KeyValuePair<OfficeXmlNode, bool> candidate in rootOmissions) {
            yield return new WordStoryRootCandidate(candidate.Key, candidate.Value);
        }
    }

    private static void RecordStoryReference<TKey>(Dictionary<TKey, bool> references, TKey id, bool omitted)
        where TKey : notnull {
        if (!references.TryGetValue(id, out bool existing) || existing && !omitted) references[id] = omitted;
    }

    private static bool RecordStoryRoot(
        Dictionary<OfficeXmlNode, bool> roots,
        OfficeXmlNode root,
        bool omitted) {
        if (!roots.TryGetValue(root, out bool existing)) {
            roots.Add(root, omitted);
            return true;
        }
        if (existing && !omitted) {
            roots[root] = false;
            return true;
        }
        return false;
    }

    private readonly struct WordStoryRootCandidate {
        internal WordStoryRootCandidate(OfficeXmlNode root, bool omittedByConverter) {
            Root = root;
            OmittedByConverter = omittedByConverter;
        }

        internal OfficeXmlNode Root { get; }

        internal bool OmittedByConverter { get; }
    }

    private static IEnumerable<(OfficeXmlNode Run, bool OmittedByConverter)> EnumerateConvertibleWordRuns(
        OfficeXmlNode container,
        bool nestedRevisionIsOmitted = false,
        bool omittedByConverter = false) {
        foreach (OfficeXmlNode child in container.Elements()) {
            string kind = child.NamespaceId == WordXml.W && child.IsModelElement() ? child.LocalName : string.Empty;
            switch (kind) {
                case "r":
                    yield return (child, omittedByConverter);
                    break;
                case "fldSimple":
                    foreach ((OfficeXmlNode Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                case "ins":
                case "moveTo":
                    foreach ((OfficeXmlNode Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                case "del":
                case "moveFrom":
                    foreach ((OfficeXmlNode Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter || nestedRevisionIsOmitted)) {
                        yield return nested;
                    }
                    break;
                case "hyperlink":
                case "sdt":
                case "sdtContent":
                    foreach ((OfficeXmlNode Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                default:
                    foreach (OfficeXmlNode omittedRun in child.ModelDescendants("r")) {
                        yield return (omittedRun, true);
                    }
                    break;
            }
        }
    }

    private static IEnumerable<(OfficeXmlNode Run, bool OmittedByConverter)> EnumerateConvertibleWordRunsWithFieldState(
        OfficeXmlNode container) {
        List<(OfficeXmlNode Run, bool OmittedByConverter)> candidates = EnumerateConvertibleWordRuns(container).ToList();
        var omittedByFieldState = new bool[candidates.Count];
        var complexFieldResults = new List<bool>();
        var complexFieldStarts = new List<int>();
        int fieldsAwaitingResult = 0;
        for (int index = 0; index < candidates.Count; index++) {
            (OfficeXmlNode Run, bool OmittedByConverter) candidate = candidates[index];
            OfficeXmlNode run = candidate.Run;
            string? markerType = run.WChild("fldChar").WAttr("fldCharType");
            bool omitted = candidate.OmittedByConverter;
            if (markerType == "begin") {
                omitted = true;
                complexFieldResults.Add(false);
                complexFieldStarts.Add(index);
                fieldsAwaitingResult++;
            } else if (complexFieldResults.Count > 0) {
                omitted |= fieldsAwaitingResult > 0;
                if (markerType == "separate") {
                    omitted = true;
                    int last = complexFieldResults.Count - 1;
                    if (!complexFieldResults[last]) {
                        complexFieldResults[last] = true;
                        fieldsAwaitingResult--;
                    }
                } else if (markerType == "end") {
                    omitted = true;
                    int last = complexFieldResults.Count - 1;
                    if (!complexFieldResults[last]) fieldsAwaitingResult--;
                    complexFieldResults.RemoveAt(last);
                    complexFieldStarts.RemoveAt(last);
                }
            }
            omittedByFieldState[index] = omitted;
        }

        // Unterminated complex fields are never completed into the RTF paragraph, so their
        // accumulated result runs are omitted even when a Separate marker was present.
        if (complexFieldStarts.Count > 0) {
            int firstUnterminatedField = complexFieldStarts[0];
            for (int index = firstUnterminatedField; index < omittedByFieldState.Length; index++) {
                omittedByFieldState[index] = true;
            }
        }

        for (int index = 0; index < candidates.Count; index++) {
            yield return (candidates[index].Run, omittedByFieldState[index]);
        }
    }

    private static IEnumerable<WordElement> EnumerateHeaderFooterElements(WordDocument document) {
        var visited = new HashSet<WordHeaderFooter>();
        foreach (WordSection section in document.Sections) {
            WordHeaderFooter?[] stories = {
                section.PageSetup.GetHeader(WordHeaderFooterType.Default), section.PageSetup.GetHeader(WordHeaderFooterType.First), section.PageSetup.GetHeader(WordHeaderFooterType.Even),
                section.PageSetup.GetFooter(WordHeaderFooterType.Default), section.PageSetup.GetFooter(WordHeaderFooterType.First), section.PageSetup.GetFooter(WordHeaderFooterType.Even)
            };
            foreach (WordHeaderFooter? story in stories) {
                if (story == null || !visited.Add(story)) continue;
                foreach (WordElement element in EnumerateWordElements(story.Elements)) yield return element;
            }
        }
    }

    private static IEnumerable<WordElement> EnumerateNoteElements(
        WordDocument document,
        ISet<long> footnoteIds,
        ISet<long> endnoteIds) {
        var visited = new HashSet<string>(StringComparer.Ordinal);
        foreach (WordFootNote note in document.FootNotes) {
            if (note.ReferenceId is not long id || !footnoteIds.Contains(id)) continue;
            string key = "F:" + (note.ReferenceId?.ToString() ?? "unknown");
            if (!visited.Add(key) || note.Paragraphs == null) continue;
            foreach (WordElement element in EnumerateWordElements(note.Paragraphs)) yield return element;
        }
        foreach (WordEndNote note in document.EndNotes) {
            if (note.ReferenceId is not long id || !endnoteIds.Contains(id)) continue;
            string key = "E:" + (note.ReferenceId?.ToString() ?? "unknown");
            if (!visited.Add(key) || note.Paragraphs == null) continue;
            foreach (WordElement element in EnumerateWordElements(note.Paragraphs)) yield return element;
        }
    }

    private static IEnumerable<WordElement> EnumerateCommentElements(
        WordDocument document,
        ISet<string> commentIds) {
        var visited = new HashSet<WordComment>();
        foreach (WordComment comment in document.Comments) {
            if (comment.Id == null || !commentIds.Contains(comment.Id) || !visited.Add(comment)) continue;
            foreach (WordElement element in EnumerateWordElements(comment.Paragraphs)) yield return element;
        }
    }

    private static bool IsUnsupportedWordElement(WordElement element) =>
        element is WordShape
        || element is WordEmbeddedDocument
        || element is WordChart
        || element is WordSmartArt
        || element is WordTextBox
        || element is WordStructuredDocumentTag;
}
