using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeParagraphs(WordDocument source, WordDocument target, WordComparisonResult result) {
            List<ParagraphSnapshot> sourceParagraphs = GetLogicalBodyParagraphs(source);
            List<ParagraphSnapshot> targetParagraphs = GetLogicalBodyParagraphs(target);
            IReadOnlyList<MatchedIndexPair> matchedParagraphs = FindMatchingIndexes(
                sourceParagraphs,
                targetParagraphs,
                ParagraphSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedParagraphs) {
                AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, sourceParagraphs.Count, targetStart, targetParagraphs.Count, result);
        }

        private static void AddParagraphRangeFindings(
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetTextSimilarity(sourceParagraphs[sourceIndex].Text, targetParagraphs[targetIndex + 1].Text) >
                    GetTextSimilarity(sourceParagraphs[sourceIndex].Text, targetParagraphs[targetIndex].Text)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Inserted,
                        ParagraphLocation(targetIndex),
                        null,
                        targetIndex,
                        null,
                        targetParagraphs[targetIndex].Text,
                        "Paragraph inserted."));
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetTextSimilarity(sourceParagraphs[sourceIndex + 1].Text, targetParagraphs[targetIndex].Text) >
                    GetTextSimilarity(sourceParagraphs[sourceIndex].Text, targetParagraphs[targetIndex].Text)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Deleted,
                        ParagraphLocation(sourceIndex),
                        sourceIndex,
                        null,
                        sourceParagraphs[sourceIndex].Text,
                        null,
                        "Paragraph deleted."));
                    sourceIndex++;
                    continue;
                }

                string sourceText = sourceParagraphs[sourceIndex].Text;
                string targetText = targetParagraphs[targetIndex].Text;

                if (!string.Equals(sourceParagraphs[sourceIndex].PartKind, targetParagraphs[targetIndex].PartKind, StringComparison.Ordinal) &&
                    string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Deleted,
                        ParagraphLocation(sourceIndex),
                        sourceIndex,
                        null,
                        sourceText,
                        null,
                        "Paragraph deleted."));
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Inserted,
                        ParagraphLocation(targetIndex),
                        null,
                        targetIndex,
                        null,
                        targetText,
                        "Paragraph inserted."));
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                if (!string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Modified,
                        ParagraphLocation(targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        "Paragraph text changed."));
                }

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Inserted,
                    ParagraphLocation(targetIndex),
                    null,
                    targetIndex,
                    null,
                    targetParagraphs[targetIndex].Text,
                    "Paragraph inserted."));
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Paragraph,
                    WordComparisonChangeKind.Deleted,
                    ParagraphLocation(sourceIndex),
                    sourceIndex,
                    null,
                    sourceParagraphs[sourceIndex].Text,
                    null,
                    "Paragraph deleted."));
                sourceIndex++;
            }
        }

        private static List<ParagraphSnapshot> GetLogicalBodyParagraphs(WordDocument document) {
            var snapshots = new List<ParagraphSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddParagraphSnapshots(snapshots, mainPart?.Document?.Body, "body");

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddParagraphSnapshots(snapshots, headerPart.Header, "header");
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddParagraphSnapshots(snapshots, footerPart.Footer, "footer");
                }
            }

            return snapshots;
        }

        private static void AddParagraphSnapshots(List<ParagraphSnapshot> snapshots, OpenXmlElement? container, string partKind) {
            IEnumerable<Paragraph> paragraphs = container?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>();
            foreach (Paragraph paragraph in paragraphs) {
                if (paragraph.Ancestors<TableCell>().Any()) {
                    continue;
                }

                string text = GetParagraphText(paragraph);
                if (text.Length == 0 && HasImageContent(paragraph)) {
                    continue;
                }

                snapshots.Add(new ParagraphSnapshot(partKind, text));
            }
        }

        private static string GetParagraphText(Paragraph paragraph) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement element in paragraph.Descendants()) {
                switch (element) {
                    case Text text:
                        builder.Append(text.Text);
                        break;
                    case TabChar:
                        builder.Append('\t');
                        break;
                    case Break breakNode:
                        if (breakNode.Type == null || breakNode.Type.Value == BreakValues.TextWrapping) {
                            builder.Append('\n');
                        } else if (breakNode.Type.Value == BreakValues.Page) {
                            builder.Append("[PageBreak]");
                        } else if (breakNode.Type.Value == BreakValues.Column) {
                            builder.Append("[ColumnBreak]");
                        } else {
                            builder.Append("[Break:");
                            builder.Append(breakNode.Type.Value.ToString());
                            builder.Append(']');
                        }

                        break;
                }
            }

            return builder.ToString();
        }

        private static bool HasImageContent(Paragraph paragraph) {
            return paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
                   paragraph.Descendants<V.ImageData>().Any();
        }

        private static double GetTextSimilarity(string source, string target) {
            if (string.Equals(source, target, StringComparison.Ordinal)) {
                return 1;
            }

            if (source.Length == 0 || target.Length == 0) {
                return 0;
            }

            return (double)GetCommonSubsequenceLength(source, target) / Math.Max(source.Length, target.Length);
        }

        private static int GetCommonSubsequenceLength(string source, string target) {
            int[,] lengths = new int[source.Length + 1, target.Length + 1];

            for (int sourceIndex = source.Length - 1; sourceIndex >= 0; sourceIndex--) {
                for (int targetIndex = target.Length - 1; targetIndex >= 0; targetIndex--) {
                    lengths[sourceIndex, targetIndex] = source[sourceIndex] == target[targetIndex]
                        ? lengths[sourceIndex + 1, targetIndex + 1] + 1
                        : Math.Max(lengths[sourceIndex + 1, targetIndex], lengths[sourceIndex, targetIndex + 1]);
                }
            }

            return lengths[0, 0];
        }

        private static string ParagraphLocation(int paragraphIndex) {
            return "paragraph[" + paragraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private sealed class ParagraphSnapshot {
            internal ParagraphSnapshot(string partKind, string text) {
                PartKind = partKind;
                Text = text;
            }

            internal string PartKind { get; }

            internal string Text { get; }
        }

        private sealed class ParagraphSnapshotEqualityComparer : IEqualityComparer<ParagraphSnapshot> {
            internal static readonly ParagraphSnapshotEqualityComparer Instance = new();

            public bool Equals(ParagraphSnapshot? x, ParagraphSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return string.Equals(x.PartKind, y.PartKind, StringComparison.Ordinal) &&
                       string.Equals(x.Text, y.Text, StringComparison.Ordinal);
            }

            public int GetHashCode(ParagraphSnapshot obj) {
                unchecked {
                    return (StringComparer.Ordinal.GetHashCode(obj.PartKind) * 397) ^
                           StringComparer.Ordinal.GetHashCode(obj.Text);
                }
            }
        }
    }
}
