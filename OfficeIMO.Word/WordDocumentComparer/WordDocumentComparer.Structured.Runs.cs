using OfficeIMO.Word.OpenXml;

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
                AnalyzeResegmentedRunFormatting(sourceRuns, targetRuns, sourceParagraphIndex, targetParagraphIndex, result, options);
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

        private static void AnalyzeResegmentedRunFormatting(
            IReadOnlyList<RunSnapshot> sourceRuns,
            IReadOnlyList<RunSnapshot> targetRuns,
            int sourceParagraphIndex,
            int targetParagraphIndex,
            WordComparisonResult result,
            WordComparisonOptions options) {
            if (!options.CompareRunFormatting) {
                return;
            }

            List<RunTextSegment> sourceSegments = CreateRunTextSegments(sourceRuns);
            List<RunTextSegment> targetSegments = CreateRunTextSegments(targetRuns);
            foreach (RunTextSegment targetSegment in targetSegments) {
                if (targetSegment.Length == 0) {
                    continue;
                }

                bool overlaps = false;
                int? sourceRunIndex = null;
                foreach (RunTextSegment sourceSegment in sourceSegments) {
                    if (!sourceSegment.Overlaps(targetSegment)) {
                        continue;
                    }

                    overlaps = true;
                    if (!string.Equals(sourceSegment.FormatSignature, targetSegment.FormatSignature, StringComparison.Ordinal)) {
                        sourceRunIndex = sourceSegment.Run.Index;
                        break;
                    }
                }

                if (!overlaps || sourceRunIndex == null) {
                    continue;
                }

                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Run,
                    WordComparisonChangeKind.Modified,
                    RunLocation(targetParagraphIndex, targetSegment.Run.Index),
                    sourceRunIndex.Value,
                    targetSegment.Run.Index,
                    targetSegment.Run.Text,
                    targetSegment.Run.Text,
                    "Run formatting changed.",
                    RunLocation(targetParagraphIndex, targetSegment.Run.Index) + "/sourceParagraph[" + sourceParagraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]"),
                    targetSegment.Run.DocumentOrder);
            }
        }

        private static List<RunTextSegment> CreateRunTextSegments(IReadOnlyList<RunSnapshot> runs) {
            var segments = new List<RunTextSegment>();
            int offset = 0;
            foreach (RunSnapshot run in runs) {
                int length = run.ComparisonText.Length;
                segments.Add(new RunTextSegment(run, offset, offset + length));
                offset += length;
            }

            return segments;
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
            foreach (OfficeXmlNode run in EnumerateComparableDescendants(paragraph.Paragraph)) {
                if (!run.Is(W, "r") || AncestorW(run, "p") != paragraph.Paragraph) {
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

        private static string GetRunText(OfficeXmlNode run) {
            var parts = new List<string>();
            foreach (OfficeXmlNode element in EnumerateComparableDescendants(run)) {
                if (element.NamespaceId != W) {
                    continue;
                }

                switch (element.LocalName) {
                    case "t":
                    case "delText":
                        parts.Add(element.InnerText);
                        break;
                    case "tab" when IsTabChar(element):
                        parts.Add("\t");
                        break;
                    case "br":
                        parts.Add(GetBreakText(element));
                        break;
                    case "sym":
                        parts.Add("[Symbol:" + (element.Attribute(W, "font") ?? string.Empty) + ":" + (element.Attribute(W, "char") ?? string.Empty) + "]");
                        break;
                    case "noBreakHyphen":
                        parts.Add("-");
                        break;
                    case "softHyphen":
                        parts.Add("[SoftHyphen]");
                        break;
                    case "cr":
                        parts.Add("\n");
                        break;
                }
            }

            return string.Concat(parts);
        }

        private static string GetBreakText(OfficeXmlNode breakNode) {
            switch (GetBreakType(breakNode)) {
                case "page":
                    return "[PageBreak]";
                case "column":
                    return "[ColumnBreak]";
                default:
                    return "\n";
            }
        }

        private static string GetRunFormatSignature(OfficeXmlNode run, WordPartView? part, OfficeXmlNode paragraph, WordComparisonOptions options) {
            if (options.CompareEffectiveFormatting) {
                return GetEffectiveRunFormatSignature(run, paragraph, part, options);
            }

            OfficeXmlNode properties = run.Child(W, "rPr");
            if (properties.IsNone) {
                return string.Empty;
            }

            return GetElementSignature(properties, options.CompareRunStyleIds ? RunFormatSignatureRules.WithStyle : RunFormatSignatureRules.WithoutStyle);
        }

        /// <summary>Run properties without their revision record, and without the run style when styles aren't compared.</summary>
        private sealed class RunFormatSignatureRules : SignatureRules {
            internal static readonly RunFormatSignatureRules WithStyle = new(dropStyle: false);
            internal static readonly RunFormatSignatureRules WithoutStyle = new(dropStyle: true);
            private readonly bool _dropStyle;

            private RunFormatSignatureRules(bool dropStyle) => _dropStyle = dropStyle;

            internal override bool Skip(OfficeXmlNode element) => element.Is(W, "rPrChange") || _dropStyle && element.Is(W, "rStyle");
        }

        private static string GetRunMatchKey(string text, string formatSignature) {
            return text + "\u001f" + formatSignature;
        }

        private static string RunLocation(int paragraphIndex, int runIndex) {
            return ParagraphLocation(paragraphIndex) + "/run[" + runIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private sealed class RunSnapshot : IComparisonFingerprint {
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

            public ulong ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }

        private readonly struct RunTextSegment {
            internal RunTextSegment(RunSnapshot run, int start, int end) {
                Run = run;
                Start = start;
                End = end;
            }

            internal RunSnapshot Run { get; }

            internal int Start { get; }

            internal int End { get; }

            internal int Length => End - Start;

            internal string FormatSignature => Run.FormatSignature;

            internal bool Overlaps(RunTextSegment other) {
                return Start < other.End && other.Start < End;
            }
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
