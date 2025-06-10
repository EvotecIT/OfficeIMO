using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

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

    private Numbering GetNumbering() {
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
                string rPr1 = "";
                string rPr2 = "";
                if (runs[i].RunProperties != null) rPr1 = runs[i].RunProperties.OuterXml;
                if (runs[i + 1].RunProperties != null) rPr2 = runs[i + 1].RunProperties.OuterXml;
                if (rPr1 == rPr2) {
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
}
