namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one decoded SxFilt PivotTable rule-filter entry.
    /// </summary>
    public sealed class LegacyXlsPivotRuleFilter {
        /// <summary>
        /// Creates decoded PivotTable rule-filter metadata.
        /// </summary>
        public LegacyXlsPivotRuleFilter(
            ushort axisFlags,
            short fieldPosition,
            short fieldReferenceIndex,
            bool selected,
            ushort subtotalFlags,
            ushort itemIndexCount) {
            AxisFlags = axisFlags;
            AxisName = GetAxisName(axisFlags);
            FieldPosition = fieldPosition;
            FieldReferenceIndex = fieldReferenceIndex;
            FieldReferenceName = GetFieldReferenceName(fieldReferenceIndex);
            Selected = selected;
            SubtotalFlags = subtotalFlags;
            SubtotalFunctionNames = GetSubtotalFunctionNames(subtotalFlags);
            ItemIndexCount = itemIndexCount;
        }

        /// <summary>Gets the raw SxFilt axis flag bits.</summary>
        public ushort AxisFlags { get; }

        /// <summary>Gets the decoded PivotTable axis name.</summary>
        public string AxisName { get; }

        /// <summary>Gets the zero-based field position within the decoded axis.</summary>
        public short FieldPosition { get; }

        /// <summary>Gets the decoded pivot, cache, or data-field reference index.</summary>
        public short FieldReferenceIndex { get; }

        /// <summary>Gets a stable field-reference name.</summary>
        public string FieldReferenceName { get; }

        /// <summary>Gets whether this filter includes the referenced field header.</summary>
        public bool Selected { get; }

        /// <summary>Gets raw subtotal-function flags from SxFilt.</summary>
        public ushort SubtotalFlags { get; }

        /// <summary>Gets decoded subtotal-function names from SxFilt.</summary>
        public IReadOnlyList<string> SubtotalFunctionNames { get; }

        /// <summary>Gets the number of SxItm indexes following this filter entry.</summary>
        public ushort ItemIndexCount { get; }

        private static string GetAxisName(ushort value) {
            bool row = (value & 0x0001) != 0;
            bool column = (value & 0x0002) != 0;
            bool page = (value & 0x0004) != 0;
            bool data = (value & 0x0008) != 0;
            if (!row && !column && !page && !data) {
                return "None";
            }

            var names = new List<string>(4);
            if (row) {
                names.Add("Row");
            }

            if (column) {
                names.Add("Column");
            }

            if (page) {
                names.Add("Page");
            }

            if (data) {
                names.Add("Data");
            }

            return string.Join("+", names);
        }

        private static string GetFieldReferenceName(short value) {
            if (value == -2) {
                return "DataField";
            }

            if (value == -1) {
                return "NoFieldReference";
            }

            return $"FieldIndex:{value}";
        }

        private static IReadOnlyList<string> GetSubtotalFunctionNames(ushort value) {
            var names = new List<string>(13);
            AddSubtotalName(names, value, 0x0001, "Data");
            AddSubtotalName(names, value, 0x0002, "Default");
            AddSubtotalName(names, value, 0x0004, "Sum");
            AddSubtotalName(names, value, 0x0008, "CountA");
            AddSubtotalName(names, value, 0x0010, "Average");
            AddSubtotalName(names, value, 0x0020, "Max");
            AddSubtotalName(names, value, 0x0040, "Min");
            AddSubtotalName(names, value, 0x0080, "Product");
            AddSubtotalName(names, value, 0x0100, "Count");
            AddSubtotalName(names, value, 0x0200, "StdDev");
            AddSubtotalName(names, value, 0x0400, "StdDevPopulation");
            AddSubtotalName(names, value, 0x0800, "Variance");
            AddSubtotalName(names, value, 0x1000, "VariancePopulation");
            AddSubtotalName(names, value, 0x4000, "Blank");
            return names;
        }

        private static void AddSubtotalName(List<string> names, ushort value, ushort flag, string name) {
            if ((value & flag) != 0) {
                names.Add(name);
            }
        }
    }
}
