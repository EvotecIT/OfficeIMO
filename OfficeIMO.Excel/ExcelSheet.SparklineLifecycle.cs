using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    /// <summary>Format-neutral snapshot of one worksheet sparkline.</summary>
    public sealed class ExcelSparklineInfo {
        internal ExcelSparklineInfo(int groupIndex, string dataRange, string locationRange, ExcelSparklineType type) {
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
        public ExcelSparklineType Type { get; }
    }

    public partial class ExcelSheet {
        /// <summary>Lists authored worksheet sparklines in package order.</summary>
        public IReadOnlyList<ExcelSparklineInfo> GetSparklines() =>
            ReadWorksheetObject(worksheet => {
                var result = new List<ExcelSparklineInfo>();
                int groupIndex = 0;
                foreach (OfficeXmlNode group in FindSparklineGroups(worksheet)) {
                    ExcelSparklineType type = ReadSparklineType(group);
                    foreach (OfficeXmlNode sparkline in FindSparklines(group)) {
                        result.Add(new ExcelSparklineInfo(
                            groupIndex,
                            sparkline.Child(ExcelFormulaNamespace, "f").InnerText,
                            sparkline.Child(ExcelFormulaNamespace, "sqref").InnerText,
                            type));
                    }
                    groupIndex++;
                }
                return new ReadOnlyCollection<ExcelSparklineInfo>(result);
            });

        /// <summary>Changes the type of every sparkline whose destination intersects a range.</summary>
        public int SetSparklineType(string locationRange, ExcelSparklineType type) {
            string typeName = SparklineTypeName(type);
            ExcelReference target = ParseLocalSparklineReference(locationRange);
            int changed = 0;
            EditWorksheetObject(worksheet => {
                foreach (OfficeXmlNode group in FindSparklineGroups(worksheet).ToList()) {
                    if (ReadSparklineType(group) == type) continue;
                    OfficeXmlNode[] all = FindSparklines(group).ToArray();
                    OfficeXmlNode[] matches = all.Where(sparkline => SparklineLocationIntersects(sparkline, target)).ToArray();
                    if (matches.Length == 0) continue;
                    if (matches.Length == all.Length) {
                        group.SetAttribute("type", typeName);
                    } else {
                        OfficeXmlNode changedGroup = group.Clone();
                        foreach (OfficeXmlNode sparkline in FindSparklines(changedGroup).ToList()) {
                            if (!SparklineLocationIntersects(sparkline, target)) sparkline.Remove();
                        }
                        foreach (OfficeXmlNode match in matches) match.Remove();
                        changedGroup.SetAttribute("type", typeName);
                        group.InsertAfterSelf(changedGroup);
                    }
                    changed += matches.Length;
                }
                return changed != 0;
            }, save: true);
            return changed;
        }

        /// <summary>Removes every sparkline whose destination intersects a range and prunes empty groups.</summary>
        public int RemoveSparklines(string locationRange) {
            ExcelReference target = ParseLocalSparklineReference(locationRange);
            return RemoveSparklinesCore(sparkline => SparklineLocationIntersects(sparkline, target));
        }

        /// <summary>Removes all worksheet sparkline groups.</summary>
        public int ClearSparklines() => RemoveSparklinesCore(static _ => true);

        private int RemoveSparklinesCore(Func<OfficeXmlNode, bool> remove) {
            int removed = 0;
            EditWorksheetObject(worksheet => {
                foreach (OfficeXmlNode group in FindSparklineGroups(worksheet).ToList()) {
                    foreach (OfficeXmlNode sparkline in FindSparklines(group).ToList()) {
                        if (!remove(sparkline)) continue;
                        sparkline.Remove();
                        removed++;
                    }
                    if (FindSparklines(group).Any()) continue;
                    OfficeXmlNode collection = group.Parent;
                    OfficeXmlNode extension = collection.Parent;
                    OfficeXmlNode list = extension.Parent;
                    group.Remove();
                    if (collection.FirstChild.IsNone) collection.Remove();
                    if (extension.FirstChild.IsNone) extension.Remove();
                    if (list.FirstChild.IsNone) list.Remove();
                }
                return removed != 0;
            }, save: true);
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

        private static bool SparklineLocationIntersects(OfficeXmlNode sparkline, ExcelReference target) {
            string? text = sparkline.Child(ExcelFormulaNamespace, "sqref").InnerText;
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
