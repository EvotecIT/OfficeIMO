using System.Collections.Generic;
using System.Linq;
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
        IEnumerable<IRtfBlock> blocks = document.Sections.Count > 0
            ? document.Sections.SelectMany(section => section.Blocks)
            : document.Blocks;
        foreach (IRtfBlock block in blocks) {
            CountUnsupportedRtfBlock(block, ref objectCount, ref shapeCount);
        }
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                CountUnsupportedRtfBlock(paragraph, ref objectCount, ref shapeCount);
            }
        }

        AddUnsupportedRtfDiagnostics(report, objectCount, shapeCount);
    }

    private static void CountUnsupportedRtfBlock(IRtfBlock block, ref int objectCount, ref int shapeCount) {
        switch (block) {
            case RtfObject:
                objectCount++;
                break;
            case RtfShape shape:
                shapeCount++;
                foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) {
                    CountUnsupportedRtfInlines(paragraph.Inlines, ref objectCount, ref shapeCount);
                }
                break;
            case RtfParagraph paragraph:
                CountUnsupportedRtfInlines(paragraph.Inlines, ref objectCount, ref shapeCount);
                break;
            case RtfTable table:
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        foreach (IRtfBlock child in cell.Blocks) {
                            CountUnsupportedRtfBlock(child, ref objectCount, ref shapeCount);
                        }
                    }
                }
                break;
        }
    }

    private static void CountUnsupportedRtfInlines(IReadOnlyList<IRtfInline> inlines, ref int objectCount, ref int shapeCount) {
        foreach (IRtfInline inline in inlines) {
            switch (inline) {
                case RtfObject:
                    objectCount++;
                    break;
                case RtfShape:
                    shapeCount++;
                    break;
                case RtfField field:
                    CountUnsupportedRtfInlines(field.Result.Inlines, ref objectCount, ref shapeCount);
                    break;
            }
        }
    }

    private static void AddUnsupportedRtfDiagnostics(RtfConversionReport report, int objectCount, int shapeCount) {
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
    }

    private static void AddWordToRtfDiagnostics(WordDocument document, RtfConversionReport report) {
        int equationCount = EnumerateWordElements(document.Elements)
            .Concat(EnumerateHeaderFooterElements(document))
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

        var unsupported = EnumerateWordElements(document.Elements)
            .Concat(EnumerateHeaderFooterElements(document))
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

    private static bool IsUnsupportedWordElement(WordElement element) =>
        element is WordShape
        || element is WordEmbeddedDocument
        || element is WordChart
        || element is WordSmartArt
        || element is WordTextBox
        || element is WordStructuredDocumentTag;
}
