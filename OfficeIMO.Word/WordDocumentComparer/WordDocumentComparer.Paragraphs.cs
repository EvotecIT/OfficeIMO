using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void CompareParagraphs(WordDocument source, WordDocument target, WordDocument result) {
            int count = Math.Min(source.Paragraphs.Count, target.Paragraphs.Count);

            for (int i = 0; i < count; i++) {
                CompareParagraph(source.Paragraphs[i], target.Paragraphs[i], result.Paragraphs[i]);
            }

            for (int i = count; i < target.Paragraphs.Count; i++) {
                WordParagraph p = result.AddParagraph();
                p.AddInsertedText(target.Paragraphs[i].Text, "Comparer");
            }

            for (int i = count; i < source.Paragraphs.Count; i++) {
                WordParagraph p = result.AddParagraph();
                p.AddDeletedText(source.Paragraphs[i].Text, "Comparer");
            }
        }

        private static void CompareParagraph(WordParagraph source, WordParagraph target, WordParagraph result) {
            var srcRuns = source._paragraph.Elements<Run>().ToList();
            var tgtRuns = target._paragraph.Elements<Run>().ToList();
            int runCount = Math.Min(srcRuns.Count, tgtRuns.Count);

            result._paragraph.RemoveAllChildren();

            for (int i = 0; i < runCount; i++) {
                string srcText = srcRuns[i].InnerText;
                string tgtText = tgtRuns[i].InnerText;

                string common = GetCommonPrefix(srcText, tgtText);
                if (!string.IsNullOrEmpty(common)) {
                    result.AddText(common);
                }

                string deleted = srcText.Substring(common.Length);
                if (!string.IsNullOrEmpty(deleted)) {
                    result.AddDeletedText(deleted, "Comparer");
                }

                string inserted = tgtText.Substring(common.Length);
                if (!string.IsNullOrEmpty(inserted)) {
                    result.AddInsertedText(inserted, "Comparer");
                }

                Run? resRun = result._paragraph.Elements<Run>().LastOrDefault();
                ApplyFormattingChange(srcRuns[i], tgtRuns[i], resRun);
            }

            for (int i = runCount; i < tgtRuns.Count; i++) {
                result.AddInsertedText(tgtRuns[i].InnerText, "Comparer");
            }

            for (int i = runCount; i < srcRuns.Count; i++) {
                result.AddDeletedText(srcRuns[i].InnerText, "Comparer");
            }
        }

        private static void ApplyFormattingChange(Run srcRun, Run tgtRun, Run? resRun) {
            if (srcRun == null || tgtRun == null || resRun == null) {
                return;
            }

            string srcXml = srcRun.RunProperties?.OuterXml ?? string.Empty;
            string tgtXml = tgtRun.RunProperties?.OuterXml ?? string.Empty;

            if (srcXml != tgtXml) {
                resRun.RunProperties = tgtRun.RunProperties != null
                    ? (RunProperties)tgtRun.RunProperties.CloneNode(true)
                    : new RunProperties();
                resRun.RunProperties.RunPropertiesChange = new RunPropertiesChange() {
                    Author = "Comparer",
                    Date = DateTime.Now
                };
                var originalProps = srcRun.RunProperties != null
                    ? (RunProperties)srcRun.RunProperties.CloneNode(true)
                    : new RunProperties();
                resRun.RunProperties.RunPropertiesChange.Append(originalProps);
            }
        }

        private static string GetCommonPrefix(string a, string b) {
            int len = Math.Min(a.Length, b.Length);
            int i = 0;
            while (i < len && a[i] == b[i]) {
                i++;
            }
            return a.Substring(0, i);
        }
    }
}
