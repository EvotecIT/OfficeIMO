using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LcsMatchingUsesCachedHashesBeforeEquality() {
            MethodInfo method = typeof(WordDocumentComparer)
                .GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
                .Single(candidate => candidate.Name == "FindMatchingIndexes" && candidate.IsGenericMethodDefinition)
                .MakeGenericMethod(typeof(FingerprintProbe));
            List<FingerprintProbe> source = Enumerable.Range(0, 999).Select(index => new FingerprintProbe(index)).ToList();
            List<FingerprintProbe> target = Enumerable.Range(1_000, 999).Select(index => new FingerprintProbe(index)).ToList();
            var comparer = new ConstantHashFingerprintComparer();

            _ = method.Invoke(null, new object[] { source, target, comparer });

            Assert.Equal(1_998, comparer.HashCalls);
            Assert.Equal(0, comparer.EqualsCalls);
        }

        [Fact]
        public void BoundedTextSimilarityWorkDoesNotScaleWithInputLength() {
            MethodInfo method = typeof(WordDocumentComparer).GetMethod(
                "GetBoundedTextSimilarity",
                BindingFlags.NonPublic | BindingFlags.Static,
                null,
                new[] { typeof(string), typeof(string) },
                null)
                ?? throw new InvalidOperationException("Bounded text similarity implementation was not found.");
            var similarity = (Func<string, string, double>)method.CreateDelegate(typeof(Func<string, string, double>));

            string shortSource = "A" + new string('x', 4_094) + "B";
            string shortTarget = "C" + new string('y', 4_094) + "D";
            string longSource = "A" + new string('x', 3_999_998) + "B";
            string longTarget = "C" + new string('y', 3_999_998) + "D";

            Assert.Equal(0D, similarity(shortSource, shortTarget));
            Assert.Equal(0D, similarity(longSource, longTarget));

            const int iterations = 128;
            TimeSpan shortElapsed = MeasureSimilarity(similarity, shortSource, shortTarget, iterations);
            TimeSpan longElapsed = MeasureSimilarity(similarity, longSource, longTarget, iterations);
            TimeSpan maximumLongElapsed = TimeSpan.FromTicks((shortElapsed.Ticks * 50) + TimeSpan.FromMilliseconds(50).Ticks);

            Assert.True(
                longElapsed <= maximumLongElapsed,
                $"Bounded similarity scaled with attacker-controlled input length: short={shortElapsed.TotalMilliseconds:F1} ms, long={longElapsed.TotalMilliseconds:F1} ms, limit={maximumLongElapsed.TotalMilliseconds:F1} ms.");
        }

        [Fact]
        public void ContainmentAwareSimilarityUsesBoundedWorkForLargeText() {
            MethodInfo method = typeof(WordDocumentComparer).GetMethod(
                "GetContainmentAwareTextSimilarity",
                BindingFlags.NonPublic | BindingFlags.Static,
                null,
                new[] { typeof(string), typeof(string) },
                null)
                ?? throw new InvalidOperationException("Containment-aware text similarity implementation was not found.");
            var similarity = (Func<string, string, double>)method.CreateDelegate(typeof(Func<string, string, double>));

            string shortSource = "A" + new string('x', 4_094) + "B";
            string target = "C" + new string('y', 1_022) + "D";
            string longSource = "A" + new string('x', 3_999_998) + "B";

            const int iterations = 4_096;
            TimeSpan shortElapsed = MeasureSimilarity(similarity, shortSource, target, iterations);
            TimeSpan longElapsed = MeasureSimilarity(similarity, longSource, target, iterations);
            TimeSpan maximumLongElapsed = TimeSpan.FromTicks((shortElapsed.Ticks * 50) + TimeSpan.FromMilliseconds(50).Ticks);

            Assert.True(
                longElapsed <= maximumLongElapsed,
                $"Containment-aware similarity scaled with attacker-controlled input length: short={shortElapsed.TotalMilliseconds:F1} ms, long={longElapsed.TotalMilliseconds:F1} ms, limit={maximumLongElapsed.TotalMilliseconds:F1} ms.");
        }

        [Fact]
        public void TextSimilarityChecksTheLargeInputLimitBeforeExactEquality() {
            MethodInfo method = typeof(WordDocumentComparer).GetMethod(
                "GetTextSimilarity",
                BindingFlags.NonPublic | BindingFlags.Static,
                null,
                new[] { typeof(string), typeof(string) },
                null)
                ?? throw new InvalidOperationException("Text similarity implementation was not found.");
            var similarity = (Func<string, string, double>)method.CreateDelegate(typeof(Func<string, string, double>));

            string shortSource = new string('x', 4_095) + "A";
            string shortTarget = new string('x', 4_095) + "B";
            string longSource = new string('x', 3_999_999) + "A";
            string longTarget = new string('x', 3_999_999) + "B";

            const int iterations = 512;
            TimeSpan shortElapsed = MeasureSimilarity(similarity, shortSource, shortTarget, iterations);
            TimeSpan longElapsed = MeasureSimilarity(similarity, longSource, longTarget, iterations);
            TimeSpan maximumLongElapsed = TimeSpan.FromTicks((shortElapsed.Ticks * 50) + TimeSpan.FromMilliseconds(50).Ticks);

            Assert.True(
                longElapsed <= maximumLongElapsed,
                $"Large-input similarity performed attacker-length exact comparisons: short={shortElapsed.TotalMilliseconds:F1} ms, long={longElapsed.TotalMilliseconds:F1} ms, limit={maximumLongElapsed.TotalMilliseconds:F1} ms.");
        }

        [Fact]
        public void CompareStructureRetainsModifiedParagraphWithTwentyCharacterShift() {
            string sharedMiddle = new string(Enumerable.Range(0, 512).Select(index => (char)(0x0100 + index)).ToArray());
            string sourceText = "S" + sharedMiddle + new string('U', 20);
            string targetText = new string('T', 21) + sharedMiddle;
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_twenty_character_shift.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph(sourceText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_twenty_character_shift.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                for (int index = 0; index < 260; index++) {
                    doc.AddParagraph("Unrelated inserted candidate " + index);
                }

                doc.AddParagraph(targetText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText == targetText);
        }

        [Fact]
        public void CompareStructureConsidersInteriorCandidateBeyondFourThousandParagraphs() {
            const string sourceText = "Status pending for security review";
            const string targetText = "Status approved for security review";
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_interior_candidate.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph(sourceText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_interior_candidate.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                for (int index = 0; index < 4_096; index++) {
                    doc.AddParagraph("Unrelated inserted candidate " + index.ToString("D4"));
                }

                doc.AddParagraph(targetText);
                doc.AddParagraph("Unrelated trailing candidate");
                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText == targetText);
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText?.StartsWith("Unrelated", StringComparison.Ordinal) == true);
        }

        [Fact]
        public void CompareStructureTableSimilarityConsumesSharedComparisonWorkBudget() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_budget.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText(new string('A', 100));
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_budget.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText(new string('B', 100));
                document.Save();
            }

            var options = new WordComparisonOptions();
            using WordDocument source = WordDocument.Load(sourcePath);
            using WordDocument target = WordDocument.Load(targetPath);
            MethodInfo getTableSnapshots = typeof(WordDocumentComparer).GetMethod(
                "GetTableSnapshots",
                BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Table snapshot implementation was not found.");
            object sourceSnapshots = getTableSnapshots.Invoke(null, new object[] { source, options })
                ?? throw new InvalidOperationException("Source table snapshots were not created.");
            object targetSnapshots = getTableSnapshots.Invoke(null, new object[] { target, options })
                ?? throw new InvalidOperationException("Target table snapshots were not created.");
            object sourceSnapshot = sourceSnapshots.GetType().GetProperty("Item")?.GetValue(sourceSnapshots, new object[] { 0 })
                ?? throw new InvalidOperationException("Source table snapshot was not available.");
            object targetSnapshot = targetSnapshots.GetType().GetProperty("Item")?.GetValue(targetSnapshots, new object[] { 0 })
                ?? throw new InvalidOperationException("Target table snapshot was not available.");
            object comparisonWorkBudget = CreateComparisonWorkBudget(1);
            MethodInfo getTableSimilarity = typeof(WordDocumentComparer).GetMethod(
                "GetTableSimilarity",
                BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Table similarity implementation was not found.");

            getTableSimilarity.Invoke(null, new[] { sourceSnapshot, targetSnapshot, comparisonWorkBudget });

            Assert.Equal(0, GetRemainingComparisonWorkUnits(comparisonWorkBudget));
        }

        [Fact]
        public void CompareStructureTableCellAlignmentConsumesSharedComparisonWorkBudget() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_budget.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText(new string('A', 100));
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_budget.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText(new string('B', 100));
                table.Rows[0].Cells[1].Paragraphs[0].SetText(new string('C', 100));
                document.Save();
            }

            var options = new WordComparisonOptions();
            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);
            using WordDocument source = WordDocument.Load(sourcePath);
            using WordDocument target = WordDocument.Load(targetPath);
            object comparisonWorkBudget = CreateComparisonWorkBudget(1);
            MethodInfo analyzeTableRow = typeof(WordDocumentComparer).GetMethod(
                "AnalyzeTableRow",
                BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Table row analysis implementation was not found.");

            analyzeTableRow.Invoke(null, new object?[] {
                source.Tables[0].Rows[0],
                target.Tables[0].Rows[0],
                null,
                null,
                0,
                0,
                0,
                0,
                result,
                options,
                comparisonWorkBudget
            });

            Assert.Equal(0, GetRemainingComparisonWorkUnits(comparisonWorkBudget));
        }

        private static TimeSpan MeasureSimilarity(
            Func<string, string, double> similarity,
            string source,
            string target,
            int iterations) {
            var stopwatch = Stopwatch.StartNew();
            for (int iteration = 0; iteration < iterations; iteration++) {
                _ = similarity(source, target);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        private static object CreateComparisonWorkBudget(long maximumWorkUnits) {
            Type workBudgetType = typeof(WordDocumentComparer).GetNestedType(
                "ComparisonWorkBudget",
                BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("Comparison work budget implementation was not found.");
            ConstructorInfo constructor = workBudgetType.GetConstructor(
                BindingFlags.Instance | BindingFlags.NonPublic,
                null,
                new[] { typeof(long) },
                null)
                ?? throw new InvalidOperationException("Comparison work budget constructor was not found.");
            return constructor.Invoke(new object[] { maximumWorkUnits });
        }

        private static long GetRemainingComparisonWorkUnits(object comparisonWorkBudget) {
            FieldInfo remainingWorkUnits = comparisonWorkBudget.GetType().GetField(
                "_remainingWorkUnits",
                BindingFlags.Instance | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("Comparison work budget state was not found.");
            return (long)(remainingWorkUnits.GetValue(comparisonWorkBudget)
                ?? throw new InvalidOperationException("Comparison work budget state was not available."));
        }

        private sealed class FingerprintProbe : WordDocumentComparer.IComparisonFingerprint {
            internal FingerprintProbe(int value) {
                Value = value;
            }

            internal int Value { get; }

            public ulong ComparisonFingerprint => unchecked((ulong)Value);
        }

        private sealed class ConstantHashFingerprintComparer : IEqualityComparer<FingerprintProbe> {
            internal int EqualsCalls { get; private set; }
            internal int HashCalls { get; private set; }

            public bool Equals(FingerprintProbe? x, FingerprintProbe? y) {
                EqualsCalls++;
                return x?.Value == y?.Value;
            }

            public int GetHashCode(FingerprintProbe obj) {
                HashCalls++;
                return 1;
            }
        }
    }
}
