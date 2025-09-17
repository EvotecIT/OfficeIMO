using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Word;

/// <summary>
/// Contains internal helper methods for WordDocument.
/// </summary>
public partial class WordDocument {
    private void SaveNumbering() {
        var numbering = GetNumbering();
        if (numbering == null) {
            return;
        }

        // it seems the order of numbering instance/abstractnums in numbering matters...

        var listAbstractNum = numbering.ChildElements.OfType<AbstractNum>().ToArray();
        var listNumberingInstance = numbering.ChildElements.OfType<NumberingInstance>().ToArray();
        var listNumberPictures = numbering.ChildElements.OfType<NumberingPictureBullet>().ToArray();

        numbering.RemoveAllChildren();

        foreach (var pictureBullet in listNumberPictures) {
            numbering.Append(pictureBullet);
        }

        foreach (var abstractNum in listAbstractNum) {
            numbering.Append(abstractNum);
        }

        foreach (var numberingInstance in listNumberingInstance) {
            numbering.Append(numberingInstance);
        }
    }

    private Numbering? GetNumbering() {
        return _wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
    }

    /// <summary>
    /// Combines one or multiple Runs having same RunProperties into one Run
    /// This is very useful to fix the issue of multiple runs with same formatting
    /// that word creates without user noticing.
    /// </summary>
    /// <param name="paragraph"></param>
    private static int CombineIdenticalRuns(Paragraph paragraph) {
        List<Run> runsToRemove = new List<Run>();
        List<Run> runs = paragraph.Elements<Run>().ToList();
        for (int i = runs.Count - 2; i >= 0; i--) {
            Text text1 = runs[i].GetFirstChild<Text>();
            Text text2 = runs[i + 1].GetFirstChild<Text>();
            if (text1 != null && text2 != null) {
                if (AreRunPropertiesEqual(runs[i].RunProperties, runs[i + 1].RunProperties)) {
                    text1.Text += text2.Text;

                    // if the text doesn't have space preservation, during merge potential double spaces
                    // or start/ending spaces could be removed which will mean the view won't be so pretty
                    if (text1.Space == null) {
                        text1.Space = SpaceProcessingModeValues.Preserve;
                    }
                    runsToRemove.Add(runs[i + 1]);
                }
            }
        }

        var count = 0;
        foreach (Run run in runsToRemove) {
            run.Remove();
            count++;
        }
        return count;
    }

    private static int CleanupParagraph(Paragraph paragraph, DocumentCleanupOptions options) {
        int count = 0;

        if (options.HasFlag(DocumentCleanupOptions.RemoveEmptyRuns) || options.HasFlag(DocumentCleanupOptions.RemoveRedundantRunProperties)) {
            foreach (var run in paragraph.Elements<Run>().ToList()) {
                if (options.HasFlag(DocumentCleanupOptions.RemoveEmptyRuns) && IsRunEmpty(run)) {
                    run.Remove();
                    count++;
                    continue;
                }

                if (options.HasFlag(DocumentCleanupOptions.RemoveRedundantRunProperties) && run.RunProperties != null && !run.RunProperties.ChildElements.Any()) {
                    run.RunProperties.Remove();
                    count++;
                }
            }
        }

        if (options.HasFlag(DocumentCleanupOptions.RemoveEmptyParagraphs) && IsParagraphEmpty(paragraph)) {
            paragraph.Remove();
            return count + 1;
        }

        if (options.HasFlag(DocumentCleanupOptions.MergeIdenticalRuns)) {
            count += CombineIdenticalRuns(paragraph);
        }

        return count;
    }

    private static bool IsRunEmpty(Run run) {
        return !run.ChildElements.OfType<Text>().Any() && run.ChildElements.All(e => e is RunProperties);
    }

    private static bool IsParagraphEmpty(Paragraph paragraph) {
        return !paragraph.Elements<Run>().Any() && paragraph.ChildElements.All(e => e is ParagraphProperties);
    }

    private static bool AreRunPropertiesEqual(RunProperties? rPr1, RunProperties? rPr2) {
        if (rPr1 == null && rPr2 == null) {
            return true;
        }

        if (rPr1 == null || rPr2 == null) {
            return false;
        }

        var x1 = Canonicalize(XElement.Parse(rPr1.OuterXml));
        var x2 = Canonicalize(XElement.Parse(rPr2.OuterXml));
        return XNode.DeepEquals(x1, x2);
    }

    private static XElement Canonicalize(XElement element) {
        return new XElement(element.Name,
            element.Attributes()
                .OrderBy(a => a.Name.NamespaceName)
                .ThenBy(a => a.Name.LocalName),
            element.Elements()
                .Select(Canonicalize)
                .OrderBy(e => e.Name.NamespaceName)
                .ThenBy(e => e.Name.LocalName));
    }
}