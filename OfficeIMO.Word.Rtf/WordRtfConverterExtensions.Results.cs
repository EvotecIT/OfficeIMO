using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Word.Rtf;

/// <content>
/// Provides result-bearing Word/RTF conversion APIs and fidelity analysis.
/// </content>
public static partial class WordRtfConverterExtensions {
    /// <summary>Converts Word to RTF and reports any structure that was flattened or omitted.</summary>
    public static RtfConversionResult<RtfDocument> ToRtfDocumentResult(this WordDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfDocument converted = document.ToRtfDocument();
        var report = new RtfConversionReport();
        AddWordToRtfDiagnostics(document, report);
        return new RtfConversionResult<RtfDocument>(converted, report);
    }

    /// <summary>Converts RTF to Word and reports any structure that was flattened or omitted.</summary>
    public static RtfConversionResult<WordDocument> ToWordDocumentResult(this RtfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        WordDocument converted = document.ToWordDocument();
        var report = new RtfConversionReport();
        AddRtfToWordDiagnostics(document, report);
        return new RtfConversionResult<WordDocument>(converted, report);
    }

    /// <summary>Converts a native RTF read result to Word while preserving parser and bridge diagnostics.</summary>
    public static RtfConversionResult<WordDocument> ToWordDocumentResult(
        this RtfReadResult readResult,
        string? sourcePath = null) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        RtfConversionResult<WordDocument> converted = readResult.Document.ToWordDocumentResult();
        var report = new RtfConversionReport();
        report.AddReadDiagnostics(readResult.Diagnostics, sourcePath);
        report.Merge(converted.Report);
        return new RtfConversionResult<WordDocument>(converted.Value, report);
    }

    private static void AddRtfToWordDiagnostics(RtfDocument document, RtfConversionReport report) {
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
                "RTF images with unsupported formats or invalid payloads are not represented by the Word bridge.",
                RtfConversionAction.Omitted,
                feature: "pict",
                count: omittedImageCount);
        }
    }

    private static void AddWordToRtfDiagnostics(WordDocument document, RtfConversionReport report) {
        List<WordStoryRootCandidate> storyRoots = EnumerateConvertibleWordStoryRoots(document).ToList();
        var footnoteIds = new HashSet<long>(storyRoots
            .Select(candidate => candidate.Root)
            .OfType<Footnote>()
            .Where(note => note.Id?.Value != null)
            .Select(note => note.Id!.Value));
        var endnoteIds = new HashSet<long>(storyRoots
            .Select(candidate => candidate.Root)
            .OfType<Endnote>()
            .Where(note => note.Id?.Value != null)
            .Select(note => note.Id!.Value));
        var commentIds = new HashSet<string>(storyRoots
            .Select(candidate => candidate.Root)
            .OfType<Comment>()
            .Where(comment => !string.IsNullOrWhiteSpace(comment.Id?.Value))
            .Select(comment => comment.Id!.Value!), StringComparer.Ordinal);
        List<WordElement> elements = EnumerateWordElements(document.Elements)
            .Concat(EnumerateHeaderFooterElements(document))
            .Concat(EnumerateNoteElements(document, footnoteIds, endnoteIds))
            .Concat(EnumerateCommentElements(document, commentIds))
            .ToList();
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
        foreach ((WordImage Image, bool OmittedByConverter) candidate in EnumerateWordImageCandidates(document, storyRoots)) {
            if (candidate.OmittedByConverter) {
                omittedImageCount++;
                continue;
            }
            WordImage image = candidate.Image;
            RtfImage? converted = CreateRtfImage(image, out OfficeImageFormat sourceFormat);
            if (converted == null) {
                omittedImageCount++;
            } else if (converted.Format == RtfImageFormat.Png &&
                       sourceFormat != OfficeImageFormat.Png) {
                normalizedImageCount++;
            }
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

    private static IEnumerable<(WordImage Image, bool OmittedByConverter)> EnumerateWordImageCandidates(
        WordDocument document,
        IEnumerable<WordStoryRootCandidate> storyRoots) {
        var visitedParagraphs = new HashSet<Paragraph>();
        var visitedRuns = new HashSet<Run>();
        foreach (WordStoryRootCandidate storyRoot in storyRoots) {
            foreach (Paragraph paragraph in storyRoot.Root.Descendants<Paragraph>()) {
                if (!visitedParagraphs.Add(paragraph)) continue;
                bool paragraphOmitted = storyRoot.OmittedByConverter ||
                                        !IsParagraphConvertedFromStoryRoot(storyRoot.Root, paragraph);
                foreach ((Run Run, bool OmittedByConverter) runCandidate in EnumerateConvertibleWordRunsWithFieldState(paragraph)) {
                    if (!visitedRuns.Add(runCandidate.Run)) continue;
                    var candidate = new WordParagraph(document, paragraph, runCandidate.Run);
                    foreach (WordImage image in candidate.EnumerateImages()) {
                        yield return (image, paragraphOmitted || runCandidate.OmittedByConverter);
                    }
                }
            }
        }
    }

    private static bool IsParagraphConvertedFromStoryRoot(OpenXmlElement root, Paragraph paragraph) {
        if (root is Document) return true;
        // The Word-to-RTF bridge reads only the direct paragraph collections exposed by
        // headers, footers, notes, and comments. Nested table/text-box paragraphs are inventoried
        // for diagnostics but are not emitted by those specialized story converters.
        return ReferenceEquals(paragraph.Parent, root);
    }

    private static IEnumerable<WordStoryRootCandidate> EnumerateConvertibleWordStoryRoots(WordDocument document) {
        DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart = document.OpenXmlDocument.MainDocumentPart;
        if (mainPart == null) yield break;

        var rootOmissions = new Dictionary<OpenXmlElement, bool>();
        if (mainPart.Document != null) rootOmissions.Add(mainPart.Document, false);
        var convertedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        if (document.Sections.Count > 0) {
            foreach (HeaderReference reference in document.Sections[0]._sectionProperties.Elements<HeaderReference>()) {
                if (!string.IsNullOrWhiteSpace(reference.Id?.Value)) convertedRelationshipIds.Add(reference.Id!.Value!);
            }
            foreach (FooterReference reference in document.Sections[0]._sectionProperties.Elements<FooterReference>()) {
                if (!string.IsNullOrWhiteSpace(reference.Id?.Value)) convertedRelationshipIds.Add(reference.Id!.Value!);
            }
        }
        foreach (DocumentFormat.OpenXml.Packaging.HeaderPart part in mainPart.HeaderParts) {
            if (part.Header != null) {
                rootOmissions[part.Header] = !convertedRelationshipIds.Contains(mainPart.GetIdOfPart(part));
            }
        }
        foreach (DocumentFormat.OpenXml.Packaging.FooterPart part in mainPart.FooterParts) {
            if (part.Footer != null) {
                rootOmissions[part.Footer] = !convertedRelationshipIds.Contains(mainPart.GetIdOfPart(part));
            }
        }

        bool added;
        do {
            added = false;
            var referencedCommentIds = new Dictionary<string, bool>(StringComparer.Ordinal);
            var referencedFootnoteIds = new Dictionary<long, bool>();
            var referencedEndnoteIds = new Dictionary<long, bool>();
            foreach (KeyValuePair<OpenXmlElement, bool> candidate in rootOmissions.ToList()) {
                foreach (string id in CollectReferencedCommentIds(new[] { candidate.Key })) {
                    RecordStoryReference(referencedCommentIds, id, candidate.Value);
                }
                foreach (FootnoteReference reference in candidate.Key.Descendants<FootnoteReference>()) {
                    if (reference.Id?.Value is long id) RecordStoryReference(referencedFootnoteIds, id, candidate.Value);
                }
                foreach (EndnoteReference reference in candidate.Key.Descendants<EndnoteReference>()) {
                    if (reference.Id?.Value is long id) RecordStoryReference(referencedEndnoteIds, id, candidate.Value);
                }
            }

            if (mainPart.WordprocessingCommentsPart?.Comments != null) {
                foreach (Comment comment in mainPart.WordprocessingCommentsPart.Comments.Elements<Comment>()) {
                    if (comment.Id?.Value is string id && referencedCommentIds.TryGetValue(id, out bool omitted)) {
                        added |= RecordStoryRoot(rootOmissions, comment, omitted);
                    }
                }
            }
            if (mainPart.FootnotesPart?.Footnotes != null) {
                foreach (Footnote note in mainPart.FootnotesPart.Footnotes.Elements<Footnote>()) {
                    if (note.Id?.Value is long id && referencedFootnoteIds.TryGetValue(id, out bool omitted)) {
                        added |= RecordStoryRoot(rootOmissions, note, omitted);
                    }
                }
            }
            if (mainPart.EndnotesPart?.Endnotes != null) {
                foreach (Endnote note in mainPart.EndnotesPart.Endnotes.Elements<Endnote>()) {
                    if (note.Id?.Value is long id && referencedEndnoteIds.TryGetValue(id, out bool omitted)) {
                        added |= RecordStoryRoot(rootOmissions, note, omitted);
                    }
                }
            }
        } while (added);

        foreach (KeyValuePair<OpenXmlElement, bool> candidate in rootOmissions) {
            yield return new WordStoryRootCandidate(candidate.Key, candidate.Value);
        }
    }

    private static void RecordStoryReference<TKey>(Dictionary<TKey, bool> references, TKey id, bool omitted)
        where TKey : notnull {
        if (!references.TryGetValue(id, out bool existing) || existing && !omitted) references[id] = omitted;
    }

    private static bool RecordStoryRoot(
        Dictionary<OpenXmlElement, bool> roots,
        OpenXmlElement root,
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
        internal WordStoryRootCandidate(OpenXmlElement root, bool omittedByConverter) {
            Root = root;
            OmittedByConverter = omittedByConverter;
        }

        internal OpenXmlElement Root { get; }

        internal bool OmittedByConverter { get; }
    }

    private static IEnumerable<(Run Run, bool OmittedByConverter)> EnumerateConvertibleWordRuns(
        OpenXmlElement container,
        bool nestedRevisionIsOmitted = false,
        bool omittedByConverter = false) {
        foreach (OpenXmlElement child in container.ChildElements) {
            switch (child) {
                case Run run:
                    yield return (run, omittedByConverter);
                    break;
                case SimpleField:
                    foreach ((Run Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                case InsertedRun:
                case MoveToRun:
                    foreach ((Run Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                case DeletedRun:
                case MoveFromRun:
                    foreach ((Run Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter || nestedRevisionIsOmitted)) {
                        yield return nested;
                    }
                    break;
                case Hyperlink:
                case SdtRun:
                case SdtContentRun:
                    foreach ((Run Run, bool OmittedByConverter) nested in EnumerateConvertibleWordRuns(
                                 child,
                                 nestedRevisionIsOmitted: true,
                                 omittedByConverter: omittedByConverter)) {
                        yield return nested;
                    }
                    break;
                default:
                    foreach (Run omittedRun in child.Descendants<Run>()) {
                        yield return (omittedRun, true);
                    }
                    break;
            }
        }
    }

    private static IEnumerable<(Run Run, bool OmittedByConverter)> EnumerateConvertibleWordRunsWithFieldState(
        OpenXmlElement container) {
        List<(Run Run, bool OmittedByConverter)> candidates = EnumerateConvertibleWordRuns(container).ToList();
        var omittedByFieldState = new bool[candidates.Count];
        var complexFieldResults = new List<bool>();
        var complexFieldStarts = new List<int>();
        for (int index = 0; index < candidates.Count; index++) {
            (Run Run, bool OmittedByConverter) candidate = candidates[index];
            Run run = candidate.Run;
            FieldChar? marker = run.Elements<FieldChar>().FirstOrDefault();
            FieldCharValues? markerType = marker?.FieldCharType?.Value;
            bool omitted = candidate.OmittedByConverter;
            if (markerType == FieldCharValues.Begin) {
                omitted = true;
                complexFieldResults.Add(false);
                complexFieldStarts.Add(index);
            } else if (complexFieldResults.Count > 0) {
                omitted |= complexFieldResults.Any(capturingResult => !capturingResult);
                if (markerType == FieldCharValues.Separate) {
                    omitted = true;
                    complexFieldResults[complexFieldResults.Count - 1] = true;
                } else if (markerType == FieldCharValues.End) {
                    omitted = true;
                    complexFieldResults.RemoveAt(complexFieldResults.Count - 1);
                    complexFieldStarts.RemoveAt(complexFieldStarts.Count - 1);
                }
            }
            omittedByFieldState[index] = omitted;
        }

        // Unterminated complex fields are never completed into the RTF paragraph, so their
        // accumulated result runs are omitted even when a Separate marker was present.
        for (int frame = 0; frame < complexFieldStarts.Count; frame++) {
            for (int index = complexFieldStarts[frame]; index < omittedByFieldState.Length; index++) {
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
                section.Header.Default, section.Header.First, section.Header.Even,
                section.Footer.Default, section.Footer.First, section.Footer.Even
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
