using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeParagraphRuns(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            int sourceParagraphIndex,
            int targetParagraphIndex,
            WordComparisonResult result,
            WordComparisonOptions options) {
            List<RunSnapshot> sourceRuns = GetRunSnapshots(sourceParagraph, options);
            List<RunSnapshot> targetRuns = GetRunSnapshots(targetParagraph, options);
            if (string.Equals(sourceParagraph.ComparisonText, targetParagraph.ComparisonText, StringComparison.Ordinal) &&
                !sourceRuns.Select(run => run.ComparisonText).SequenceEqual(targetRuns.Select(run => run.ComparisonText), StringComparer.Ordinal)) {
                return;
            }

            if (sourceRuns.Count == 1 &&
                targetRuns.Count == 1 &&
                !string.Equals(sourceRuns[0].ComparisonText, targetRuns[0].ComparisonText, StringComparison.Ordinal) &&
                string.Equals(sourceRuns[0].FormatSignature, targetRuns[0].FormatSignature, StringComparison.Ordinal)) {
                return;
            }

            IReadOnlyList<MatchedIndexPair> matchedRuns = FindMatchingIndexes(
                sourceRuns,
                targetRuns,
                RunSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;
            foreach (MatchedIndexPair match in matchedRuns) {
                AddRunRangeFindings(sourceRuns, targetRuns, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, targetParagraphIndex, result);
                AnalyzeMatchedRun(sourceRuns[match.SourceIndex], targetRuns[match.TargetIndex], sourceParagraphIndex, targetParagraphIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddRunRangeFindings(sourceRuns, targetRuns, sourceStart, sourceRuns.Count, targetStart, targetRuns.Count, targetParagraphIndex, result);
        }

        private static void AddRunRangeFindings(
            IReadOnlyList<RunSnapshot> sourceRuns,
            IReadOnlyList<RunSnapshot> targetRuns,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            int targetParagraphIndex,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                AnalyzeAlignedRun(sourceRuns[sourceIndex], targetRuns[targetIndex], targetParagraphIndex, result);
                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                RunSnapshot targetRun = targetRuns[targetIndex];
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Run,
                    WordComparisonChangeKind.Inserted,
                    RunLocation(targetParagraphIndex, targetRun.Index),
                    null,
                    targetRun.Index,
                    null,
                    targetRun.Text,
                    "Run inserted."),
                    targetRun.DocumentOrder);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                RunSnapshot sourceRun = sourceRuns[sourceIndex];
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Run,
                    WordComparisonChangeKind.Deleted,
                    RunLocation(targetParagraphIndex, sourceRun.Index),
                    sourceRun.Index,
                    null,
                    sourceRun.Text,
                    null,
                    "Run deleted."),
                    sourceRun.DocumentOrder);
                sourceIndex++;
            }
        }

        private static void AnalyzeMatchedRun(
            RunSnapshot sourceRun,
            RunSnapshot targetRun,
            int sourceParagraphIndex,
            int targetParagraphIndex,
            WordComparisonResult result) {
            if (!string.Equals(sourceRun.ComparisonText, targetRun.ComparisonText, StringComparison.Ordinal)) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Run,
                    WordComparisonChangeKind.Modified,
                    RunLocation(targetParagraphIndex, targetRun.Index),
                    sourceRun.Index,
                    targetRun.Index,
                    sourceRun.Text,
                    targetRun.Text,
                    "Run text changed."),
                    targetRun.DocumentOrder);
                return;
            }

            if (!string.Equals(sourceRun.FormatSignature, targetRun.FormatSignature, StringComparison.Ordinal)) {
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Run,
                    WordComparisonChangeKind.Modified,
                    RunLocation(targetParagraphIndex, targetRun.Index),
                    sourceRun.Index,
                    targetRun.Index,
                    sourceRun.Text,
                    targetRun.Text,
                    "Run formatting changed."),
                    targetRun.DocumentOrder);
            }
        }

        private static void AnalyzeAlignedRun(
            RunSnapshot sourceRun,
            RunSnapshot targetRun,
            int targetParagraphIndex,
            WordComparisonResult result) {
            if (string.Equals(sourceRun.MatchKey, targetRun.MatchKey, StringComparison.Ordinal) &&
                string.Equals(sourceRun.ComparisonText, targetRun.ComparisonText, StringComparison.Ordinal) &&
                string.Equals(sourceRun.FormatSignature, targetRun.FormatSignature, StringComparison.Ordinal)) {
                return;
            }

            string message = string.Equals(sourceRun.ComparisonText, targetRun.ComparisonText, StringComparison.Ordinal)
                ? "Run formatting changed."
                : "Run text changed.";
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Run,
                WordComparisonChangeKind.Modified,
                RunLocation(targetParagraphIndex, targetRun.Index),
                sourceRun.Index,
                targetRun.Index,
                sourceRun.Text,
                targetRun.Text,
                message),
                targetRun.DocumentOrder);
        }

        private static List<RunSnapshot> GetRunSnapshots(ParagraphSnapshot paragraph, WordComparisonOptions options) {
            var snapshots = new List<RunSnapshot>();
            int runIndex = 0;
            foreach (Run run in EnumerateComparableDescendants(paragraph.Paragraph).OfType<Run>()) {
                if (run.Ancestors<Paragraph>().FirstOrDefault() != paragraph.Paragraph) {
                    continue;
                }

                string text = GetRunText(run);
                string comparisonText = NormalizeComparisonText(text, options);
                string formatSignature = options.CompareRunFormatting ? GetRunFormatSignature(run, paragraph.Part, paragraph.Paragraph, options) : string.Empty;
                snapshots.Add(new RunSnapshot(
                    runIndex,
                    text,
                    comparisonText,
                    GetRunMatchKey(comparisonText, formatSignature),
                    formatSignature,
                    paragraph.DocumentOrder + runIndex + 1));
                runIndex++;
            }

            return snapshots;
        }

        private static string GetRunText(Run run) {
            var parts = new List<string>();
            foreach (OpenXmlElement element in EnumerateComparableDescendants(run)) {
                switch (element) {
                    case Text text:
                        parts.Add(text.Text ?? string.Empty);
                        break;
                    case DeletedText deletedText:
                        parts.Add(deletedText.Text ?? string.Empty);
                        break;
                    case TabChar:
                        parts.Add("\t");
                        break;
                    case Break breakNode:
                        parts.Add(GetBreakText(breakNode));
                        break;
                    case SymbolChar symbol:
                        parts.Add("[Symbol:" + (symbol.Font?.Value ?? string.Empty) + ":" + (symbol.Char?.Value ?? string.Empty) + "]");
                        break;
                    case NoBreakHyphen:
                        parts.Add("-");
                        break;
                    case SoftHyphen:
                        parts.Add("[SoftHyphen]");
                        break;
                    case CarriageReturn:
                        parts.Add("\n");
                        break;
                }
            }

            return string.Concat(parts);
        }

        private static string GetBreakText(Break breakNode) {
            if (breakNode.Type == null || breakNode.Type.Value == BreakValues.TextWrapping) {
                return "\n";
            }

            if (breakNode.Type.Value == BreakValues.Page) {
                return "[PageBreak]";
            }

            if (breakNode.Type.Value == BreakValues.Column) {
                return "[ColumnBreak]";
            }

            return "[Break:" + breakNode.Type.Value + "]";
        }

        private static string GetRunFormatSignature(Run run, OpenXmlPart? part, Paragraph paragraph, WordComparisonOptions options) {
            if (options.CompareEffectiveFormatting) {
                return GetEffectiveRunFormatSignature(run, paragraph, part, options);
            }

            RunProperties? properties = run.RunProperties;
            if (properties == null) {
                return string.Empty;
            }

            OpenXmlElement clone = properties.CloneNode(true);
            foreach (RunPropertiesChange change in clone.Descendants<RunPropertiesChange>().ToList()) {
                change.Remove();
            }

            if (!options.CompareRunStyleIds) {
                foreach (RunStyle runStyle in clone.Descendants<RunStyle>().ToList()) {
                    runStyle.Remove();
                }
            }

            return clone.OuterXml;
        }

        private static string GetRunMatchKey(string text, string formatSignature) {
            return text + "\u001f" + formatSignature;
        }

        private static string RunLocation(int paragraphIndex, int runIndex) {
            return ParagraphLocation(paragraphIndex) + "/run[" + runIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private sealed class RunSnapshot {
            internal RunSnapshot(int index, string text, string comparisonText, string matchKey, string formatSignature, int documentOrder) {
                Index = index;
                Text = text;
                ComparisonText = comparisonText;
                MatchKey = matchKey;
                FormatSignature = formatSignature;
                DocumentOrder = documentOrder;
            }

            internal int Index { get; }

            internal string Text { get; }

            internal string ComparisonText { get; }

            internal string MatchKey { get; }

            internal string FormatSignature { get; }

            internal int DocumentOrder { get; }
        }

        private sealed class RunSnapshotEqualityComparer : IEqualityComparer<RunSnapshot> {
            internal static readonly RunSnapshotEqualityComparer Instance = new();

            public bool Equals(RunSnapshot? x, RunSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return string.Equals(x.MatchKey, y.MatchKey, StringComparison.Ordinal);
            }

            public int GetHashCode(RunSnapshot obj) {
                return StringComparer.Ordinal.GetHashCode(obj.MatchKey);
            }
        }
    }
}
