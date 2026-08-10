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
        List<WordElement> elements = EnumerateWordElements(document.Elements)
            .Concat(EnumerateHeaderFooterElements(document))
            .Concat(EnumerateNoteElements(document))
            .Concat(EnumerateCommentElements(document))
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
        foreach ((WordImage Image, bool OmittedByConverter) candidate in EnumerateWordImageCandidates(document)) {
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

    private static IEnumerable<(WordImage Image, bool OmittedByConverter)> EnumerateWordImageCandidates(WordDocument document) {
        var visitedParagraphs = new HashSet<Paragraph>();
        var visitedRuns = new HashSet<Run>();
        foreach (OpenXmlElement storyRoot in EnumerateConvertibleWordStoryRoots(document)) {
            foreach (Paragraph paragraph in storyRoot.Descendants<Paragraph>()) {
                if (!visitedParagraphs.Add(paragraph)) continue;
                foreach ((Run Run, bool OmittedByConverter) runCandidate in EnumerateConvertibleWordRuns(paragraph)) {
                    if (!visitedRuns.Add(runCandidate.Run)) continue;
                    var candidate = new WordParagraph(document, paragraph, runCandidate.Run);
                    foreach (WordImage image in candidate.EnumerateImages()) {
                        yield return (image, runCandidate.OmittedByConverter);
                    }
                }
            }
        }
    }

    private static IEnumerable<OpenXmlElement> EnumerateConvertibleWordStoryRoots(WordDocument document) {
        DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart = document.OpenXmlDocument.MainDocumentPart;
        if (mainPart?.Document != null) yield return mainPart.Document;
        if (mainPart == null) yield break;

        foreach (DocumentFormat.OpenXml.Packaging.HeaderPart part in mainPart.HeaderParts) {
            if (part.Header != null) yield return part.Header;
        }
        foreach (DocumentFormat.OpenXml.Packaging.FooterPart part in mainPart.FooterParts) {
            if (part.Footer != null) yield return part.Footer;
        }
        if (mainPart.FootnotesPart?.Footnotes != null) {
            yield return mainPart.FootnotesPart.Footnotes;
        }
        if (mainPart.EndnotesPart?.Endnotes != null) {
            yield return mainPart.EndnotesPart.Endnotes;
        }
        if (mainPart.WordprocessingCommentsPart?.Comments != null) {
            yield return mainPart.WordprocessingCommentsPart.Comments;
        }
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
                case SimpleField simpleField:
                    foreach (Run fieldRun in simpleField.Elements<Run>()) {
                        yield return (fieldRun, omittedByConverter);
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
            }
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

    private static IEnumerable<WordElement> EnumerateNoteElements(WordDocument document) {
        var visited = new HashSet<string>(StringComparer.Ordinal);
        foreach (WordFootNote note in document.FootNotes) {
            string key = "F:" + (note.ReferenceId?.ToString() ?? "unknown");
            if (!visited.Add(key) || note.Paragraphs == null) continue;
            foreach (WordElement element in EnumerateWordElements(note.Paragraphs)) yield return element;
        }
        foreach (WordEndNote note in document.EndNotes) {
            string key = "E:" + (note.ReferenceId?.ToString() ?? "unknown");
            if (!visited.Add(key) || note.Paragraphs == null) continue;
            foreach (WordElement element in EnumerateWordElements(note.Paragraphs)) yield return element;
        }
    }

    private static IEnumerable<WordElement> EnumerateCommentElements(WordDocument document) {
        var visited = new HashSet<WordComment>();
        foreach (WordComment comment in document.Comments) {
            if (!visited.Add(comment)) continue;
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
