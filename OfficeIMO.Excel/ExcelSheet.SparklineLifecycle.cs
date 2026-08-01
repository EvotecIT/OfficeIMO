using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;

namespace OfficeIMO.Excel {
    /// <summary>Format-neutral snapshot of one worksheet sparkline.</summary>
    public sealed class ExcelSparklineInfo {
        internal ExcelSparklineInfo(int groupIndex, string dataRange, string locationRange, SparklineTypeValues type) {
            GroupIndex = groupIndex;
            DataRange = dataRange;
            LocationRange = locationRange;
            Type = type;
        }

        /// <summary>Zero-based group index in worksheet order.</summary>
        public int GroupIndex { get; }
        /// <summary>Source data formula or range.</summary>
        public string DataRange { get; }
        /// <summary>Destination cell or range.</summary>
        public string LocationRange { get; }
        /// <summary>Sparkline rendering type.</summary>
        public SparklineTypeValues Type { get; }
    }

    public partial class ExcelSheet {
        /// <summary>Lists authored worksheet sparklines in package order.</summary>
        public IReadOnlyList<ExcelSparklineInfo> GetSparklines() {
            var result = new List<ExcelSparklineInfo>();
            int groupIndex = 0;
            foreach (SparklineGroup group in WorksheetRoot.Descendants<SparklineGroup>()) {
                SparklineTypeValues type = group.Type?.Value ?? SparklineTypeValues.Line;
                foreach (Sparkline sparkline in group.Descendants<Sparkline>()) {
                    result.Add(new ExcelSparklineInfo(
                        groupIndex,
                        sparkline.Formula?.Text ?? string.Empty,
                        sparkline.GetFirstChild<OfficeReferenceSequence>()?.Text ?? string.Empty,
                        type));
                }
                groupIndex++;
            }
            return new ReadOnlyCollection<ExcelSparklineInfo>(result);
        }

        /// <summary>Changes the type of every sparkline whose destination intersects a range.</summary>
        public int SetSparklineType(string locationRange, SparklineTypeValues type) {
            ExcelReference target = ParseLocalSparklineReference(locationRange);
            int changed = 0;
            WriteLock(() => {
                foreach (SparklineGroup group in WorksheetRoot.Descendants<SparklineGroup>().ToList()) {
                    Sparkline[] all = group.Descendants<Sparkline>().ToArray();
                    Sparkline[] matches = all.Where(sparkline => SparklineLocationIntersects(sparkline, target)).ToArray();
                    SparklineTypeValues currentType = group.Type?.Value ?? SparklineTypeValues.Line;
                    if (matches.Length == 0 || currentType == type) continue;
                    if (matches.Length == all.Length) {
                        group.Type = type;
                    } else {
                        var changedGroup = (SparklineGroup)group.CloneNode(true);
                        Sparkline[] changedClones = changedGroup.Descendants<Sparkline>().ToArray();
                        for (int index = changedClones.Length - 1; index >= 0; index--) {
                            if (!SparklineLocationIntersects(changedClones[index], target)) changedClones[index].Remove();
                        }
                        foreach (Sparkline match in matches) match.Remove();
                        changedGroup.Type = type;
                        group.Parent!.InsertAfter(changedGroup, group);
                    }
                    changed += matches.Length;
                }
                if (changed > 0) WorksheetRoot.Save();
            });
            return changed;
        }

        /// <summary>Removes every sparkline whose destination intersects a range and prunes empty groups.</summary>
        public int RemoveSparklines(string locationRange) {
            ExcelReference target = ParseLocalSparklineReference(locationRange);
            int removed = 0;
            WriteLock(() => {
                foreach (SparklineGroup group in WorksheetRoot.Descendants<SparklineGroup>().ToList()) {
                    foreach (Sparkline sparkline in group.Descendants<Sparkline>().ToList()) {
                        if (!SparklineLocationIntersects(sparkline, target)) continue;
                        sparkline.Remove();
                        removed++;
                    }
                    Sparklines? collection = group.GetFirstChild<Sparklines>();
                    if (collection == null || !collection.Elements<Sparkline>().Any()) group.Remove();
                }
                if (removed > 0) {
                    CleanupEmptySparklineStructures(WorksheetRoot);
                    WorksheetRoot.Save();
                }
            });
            return removed;
        }

        /// <summary>Removes all worksheet sparkline groups.</summary>
        public int ClearSparklines() {
            int removed = 0;
            WriteLock(() => {
                foreach (SparklineGroups groups in WorksheetRoot.Descendants<SparklineGroups>().ToList()) {
                    removed += groups.Descendants<Sparkline>().Count();
                    groups.Remove();
                }
                if (removed > 0) {
                    CleanupEmptySparklineStructures(WorksheetRoot);
                    WorksheetRoot.Save();
                }
            });
            return removed;
        }

        private ExcelReference ParseLocalSparklineReference(string reference) {
            if (!ExcelReference.TryParse(reference, out ExcelReference? parsed)) {
                throw new ArgumentException($"Invalid sparkline destination reference '{reference}'.", nameof(reference));
            }
            if (parsed!.Kind != ExcelReferenceKind.Cell && parsed.Kind != ExcelReferenceKind.Range) {
                throw new ArgumentException("Sparkline destination must be a cell or rectangular cell range.", nameof(reference));
            }
            if (parsed.IsQualified && !IsCurrentSheetQualifier(parsed.Qualifier!, Name)) {
                throw new ArgumentException("Sparkline destination must belong to the current worksheet.", nameof(reference));
            }
            return parsed!;
        }

        private static bool SparklineLocationIntersects(Sparkline sparkline, ExcelReference target) {
            string? text = sparkline.GetFirstChild<OfficeReferenceSequence>()?.Text;
            return ExcelReference.TryParse(text, out ExcelReference? location)
                && ReferencesIntersectIgnoringQualifier(location!, target);
        }

        private static bool ReferencesIntersectIgnoringQualifier(ExcelReference left, ExcelReference right) {
            left.GetBounds(out int lr1, out int lc1, out int lr2, out int lc2);
            right.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
            return lr1 <= rr2 && lr2 >= rr1 && lc1 <= rc2 && lc2 >= rc1;
        }
    }
}
