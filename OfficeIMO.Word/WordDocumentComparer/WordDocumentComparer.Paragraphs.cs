using OfficeIMO.Word.OpenXml;

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
            OfficeXmlNode sourceParagraph = WordPackageView.ElementOf(source);
            OfficeXmlNode targetParagraph = WordPackageView.ElementOf(target);
            OfficeXmlNode resultParagraph = WordPackageView.ElementOf(result);
            List<OfficeXmlNode> srcRuns = ChildrenW(sourceParagraph, "r");
            List<OfficeXmlNode> tgtRuns = ChildrenW(targetParagraph, "r");
            int runCount = Math.Min(srcRuns.Count, tgtRuns.Count);
            OfficeXmlNode paragraphProperties = resultParagraph.Child(W, "pPr");
            if (paragraphProperties.IsNone) paragraphProperties = sourceParagraph.Child(W, "pPr");
            if (paragraphProperties.IsNone) paragraphProperties = targetParagraph.Child(W, "pPr");
            paragraphProperties = paragraphProperties.IsNone ? default : paragraphProperties.Clone();

            RemoveAllChildren(resultParagraph);
            if (!paragraphProperties.IsNone) {
                resultParagraph.Append(paragraphProperties);
            }

            for (int i = 0; i < runCount; i++) {
                string srcText = InnerText(srcRuns[i]);
                string tgtText = InnerText(tgtRuns[i]);

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

                OfficeXmlNode resRun = resultParagraph.Children(W, "r").Last();
                ApplyFormattingChange(srcRuns[i], tgtRuns[i], resRun);
            }

            for (int i = runCount; i < tgtRuns.Count; i++) {
                result.AddInsertedText(InnerText(tgtRuns[i]), "Comparer");
            }

            for (int i = runCount; i < srcRuns.Count; i++) {
                result.AddDeletedText(InnerText(srcRuns[i]), "Comparer");
            }
        }

        private static bool SameBytes(byte[] left, byte[] right) {
            if (left.Length != right.Length) return false;
            for (int index = 0; index < left.Length; index++) {
                if (left[index] != right[index]) return false;
            }
            return true;
        }

        private static void ApplyFormattingChange(OfficeXmlNode srcRun, OfficeXmlNode tgtRun, OfficeXmlNode resRun) {
            if (srcRun.IsNone || tgtRun.IsNone || resRun.IsNone) {
                return;
            }

            OfficeXmlNode srcProperties = srcRun.Child(W, "rPr");
            OfficeXmlNode tgtProperties = tgtRun.Child(W, "rPr");
            if (srcProperties.IsNone == tgtProperties.IsNone &&
                (srcProperties.IsNone || SameBytes(srcProperties.OuterXml(), tgtProperties.OuterXml()))) {
                return;
            }

            OfficeXmlNode properties = tgtProperties.IsNone ? resRun.CreateElement(W, "rPr") : tgtProperties.Clone();
            properties = SetPropertyChild(resRun, properties, WRView.Order);
            OfficeXmlNode change = properties.CreateElement(W, "rPrChange");
            change.SetAttribute(W, "author", "Comparer");
            change.SetAttribute(W, "date", WordPackageView.FormatDateTime(DateTime.Now));
            change = SetPropertyChild(properties, change, order: null);
            change.Append(srcProperties.IsNone ? change.CreateElement(W, "rPr") : srcProperties.Clone());
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
