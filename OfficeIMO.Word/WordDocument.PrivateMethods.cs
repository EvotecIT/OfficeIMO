using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Word;

// Contains internal helper methods for WordDocument.
public partial class WordDocument {
    private void SaveNumbering() {
        var numbering = GetNumbering();
        if (numbering == null || HasCanonicalNumberingOrder(numbering)) {
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

    private static bool HasCanonicalNumberingOrder(Numbering numbering) {
        int previousGroup = 0;
        foreach (OpenXmlElement child in numbering.ChildElements) {
            int currentGroup = child switch {
                NumberingPictureBullet => 0,
                AbstractNum => 1,
                NumberingInstance => 2,
                _ => -1
            };
            if (currentGroup < previousGroup) {
                return false;
            }
            previousGroup = currentGroup;
        }
        return true;
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
        List<Run> runs = paragraph.Elements<Run>().ToList();
        var texts = new Text?[runs.Count];
        for (int i = 0; i < runs.Count; i++) {
            TryGetMergeableRunText(runs[i], out texts[i]);
        }

        // Each run that can join its predecessor extends a chain; the first run of the chain
        // receives the joined text once, instead of growing the text pair by pair.
        var canonicalProperties = new XElement?[runs.Count];
        var runsToRemove = new List<Run>();
        var mergedText = new StringBuilder();
        int index = 0;
        while (index < runs.Count) {
            int chainEnd = index;
            while (chainEnd + 1 < runs.Count &&
                   texts[chainEnd] != null &&
                   texts[chainEnd + 1] != null &&
                   AreRunPropertiesEqual(runs, canonicalProperties, chainEnd, chainEnd + 1)) {
                chainEnd++;
            }

            if (chainEnd > index) {
                Text first = texts[index]!;
                mergedText.Clear().Append(first.Text);
                for (int next = index + 1; next <= chainEnd; next++) {
                    mergedText.Append(texts[next]!.Text);
                    runsToRemove.Add(runs[next]);
                }
                first.Text = mergedText.ToString();

                // if the text doesn't have space preservation, during merge potential double spaces
                // or start/ending spaces could be removed which will mean the view won't be so pretty
                if (first.Space == null) {
                    first.Space = SpaceProcessingModeValues.Preserve;
                }
            }

            index = chainEnd + 1;
        }

        foreach (Run run in runsToRemove) {
            run.Remove();
        }
        return runsToRemove.Count;
    }

    private static bool TryGetMergeableRunText(Run run, out Text? text) {
        var content = run.ChildElements
            .Where(static element => element is not RunProperties)
            .ToArray();
        text = content.Length == 1 ? content[0] as Text : null;
        return text != null;
    }

    private static int CleanupParagraph(Paragraph paragraph, WordDocumentCleanupOptions options) {
        int count = 0;

        if (options.HasFlag(WordDocumentCleanupOptions.RemoveEmptyRuns) || options.HasFlag(WordDocumentCleanupOptions.RemoveRedundantRunProperties)) {
            foreach (var run in paragraph.Elements<Run>().ToList()) {
                if (options.HasFlag(WordDocumentCleanupOptions.RemoveEmptyRuns) && IsRunEmpty(run)) {
                    run.Remove();
                    count++;
                    continue;
                }

                if (options.HasFlag(WordDocumentCleanupOptions.RemoveRedundantRunProperties) && run.RunProperties != null && !run.RunProperties.ChildElements.Any()) {
                    run.RunProperties.Remove();
                    count++;
                }
            }
        }

        if (options.HasFlag(WordDocumentCleanupOptions.RemoveEmptyParagraphs) && IsParagraphEmpty(paragraph)) {
            paragraph.Remove();
            return count + 1;
        }

        if (options.HasFlag(WordDocumentCleanupOptions.MergeIdenticalRuns)) {
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

    private static bool AreRunPropertiesEqual(List<Run> runs, XElement?[] canonicalProperties, int left, int right) {
        RunProperties? rPr1 = runs[left].RunProperties;
        RunProperties? rPr2 = runs[right].RunProperties;
        if (rPr1 == null && rPr2 == null) {
            return true;
        }

        if (rPr1 == null || rPr2 == null) {
            return false;
        }

        string xml1 = rPr1.OuterXml;
        string xml2 = rPr2.OuterXml;
        if (string.Equals(xml1, xml2, StringComparison.Ordinal)) {
            return true;
        }

        // Each run is compared with both neighbours, so canonicalize it only once.
        XElement x1 = canonicalProperties[left] ??= Canonicalize(XElement.Parse(xml1));
        XElement x2 = canonicalProperties[right] ??= Canonicalize(XElement.Parse(xml2));
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
