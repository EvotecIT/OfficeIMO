using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes how a workbook feature is handled for a physical Excel format.
    /// </summary>
    public enum ExcelFormatCapabilityStatus {
        /// <summary>The feature can be authored in the target format.</summary>
        Native,
        /// <summary>A useful native subset can be authored in the target format.</summary>
        NativeSubset,
        /// <summary>The outcome is authored with compatible, broadly supported chart or workbook primitives.</summary>
        CompatibleRecipe,
        /// <summary>OfficeIMO-owned metadata can be authored and inspected, but the target application's native object is not created.</summary>
        MetadataOnly,
        /// <summary>Existing content is preserved or reported, but new native authoring is not available.</summary>
        PreserveOnly,
        /// <summary>The feature cannot be represented or safely written by the current target writer.</summary>
        Unsupported
    }

    /// <summary>
    /// One feature row in the XLSX, XLS, and XLSB capability matrix.
    /// </summary>
    public sealed class ExcelFormatCapabilityEntry {
        internal ExcelFormatCapabilityEntry(
            string feature,
            ExcelFormatCapabilityStatus xlsx,
            ExcelFormatCapabilityStatus xls,
            ExcelFormatCapabilityStatus xlsb,
            string notes) {
            Feature = feature;
            Xlsx = xlsx;
            Xls = xls;
            Xlsb = xlsb;
            Notes = notes;
        }

        /// <summary>Feature name.</summary>
        public string Feature { get; }

        /// <summary>XLSX support status.</summary>
        public ExcelFormatCapabilityStatus Xlsx { get; }

        /// <summary>BIFF8 XLS support status.</summary>
        public ExcelFormatCapabilityStatus Xls { get; }

        /// <summary>BIFF12 XLSB support status.</summary>
        public ExcelFormatCapabilityStatus Xlsb { get; }

        /// <summary>Important scope or fallback details.</summary>
        public string Notes { get; }

        /// <summary>Returns the status for a physical Excel format.</summary>
        public ExcelFormatCapabilityStatus GetStatus(ExcelFileFormat format) {
            return format switch {
                ExcelFileFormat.Xlsx => Xlsx,
                ExcelFileFormat.Xls => Xls,
                ExcelFileFormat.Xlsb => Xlsb,
                _ => throw new ArgumentOutOfRangeException(nameof(format))
            };
        }
    }

    /// <summary>
    /// Current target-format capability matrix for high-value OfficeIMO.Excel features.
    /// </summary>
    public sealed class ExcelFormatCapabilityReport {
        private ExcelFormatCapabilityReport() {
            Entries = new[] {
                new ExcelFormatCapabilityEntry(
                    "Values, dates, and common styles",
                    ExcelFormatCapabilityStatus.Native,
                    ExcelFormatCapabilityStatus.NativeSubset,
                    ExcelFormatCapabilityStatus.NativeSubset,
                    "XLS and XLSB writers cover the common cell model; advanced format-specific records remain guarded."),
                new ExcelFormatCapabilityEntry(
                    "Formula authoring and cached results",
                    ExcelFormatCapabilityStatus.Native,
                    ExcelFormatCapabilityStatus.NativeSubset,
                    ExcelFormatCapabilityStatus.NativeSubset,
                    "The lightweight evaluator and dependency diagnostics are format-neutral; BIFF token writers support a bounded formula subset."),
                new ExcelFormatCapabilityEntry(
                    "Formula dependency depth and cycle diagnostics",
                    ExcelFormatCapabilityStatus.Native,
                    ExcelFormatCapabilityStatus.Native,
                    ExcelFormatCapabilityStatus.Native,
                    "Inspection and the configurable calculation depth budget operate on the shared workbook model before target serialization."),
                new ExcelFormatCapabilityEntry(
                    "Classic chart authoring",
                    ExcelFormatCapabilityStatus.Native,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    "Native XLS and XLSB writers do not yet create drawing/chart records; existing binary chart metadata is preserved or reported where supported."),
                new ExcelFormatCapabilityEntry(
                    "Histogram, Pareto, funnel, and waterfall recipes",
                    ExcelFormatCapabilityStatus.CompatibleRecipe,
                    ExcelFormatCapabilityStatus.Unsupported,
                    ExcelFormatCapabilityStatus.Unsupported,
                    "XLSX recipes use classic chart primitives for broad Excel compatibility; true ChartEx authoring is not yet implemented."),
                new ExcelFormatCapabilityEntry(
                    "Pivot table authoring",
                    ExcelFormatCapabilityStatus.NativeSubset,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    "XLSX supports common layouts, fields, filters, grouping, calculated fields, cache records, and refresh metadata."),
                new ExcelFormatCapabilityEntry(
                    "Existing pivot source updates",
                    ExcelFormatCapabilityStatus.NativeSubset,
                    ExcelFormatCapabilityStatus.Unsupported,
                    ExcelFormatCapabilityStatus.Unsupported,
                    "XLSX source updates validate headers, guard shared caches, invalidate stale records, and request refresh on open."),
                new ExcelFormatCapabilityEntry(
                    "Slicer cache metadata",
                    ExcelFormatCapabilityStatus.MetadataOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    "OfficeIMO-owned pivot/field binding metadata can be validated and inspected; native Excel caches and UI shapes are not materialized."),
                new ExcelFormatCapabilityEntry(
                    "Timeline cache metadata",
                    ExcelFormatCapabilityStatus.MetadataOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    ExcelFormatCapabilityStatus.PreserveOnly,
                    "OfficeIMO-owned pivot/field binding metadata can be validated and inspected; native Excel caches and UI shapes are not materialized.")
            };
        }

        /// <summary>Current OfficeIMO.Excel capability matrix.</summary>
        public static ExcelFormatCapabilityReport Current { get; } = new ExcelFormatCapabilityReport();

        /// <summary>Capability rows in stable display order.</summary>
        public IReadOnlyList<ExcelFormatCapabilityEntry> Entries { get; }

        /// <summary>Finds a capability row by exact feature name.</summary>
        public ExcelFormatCapabilityEntry? Find(string feature) {
            if (string.IsNullOrWhiteSpace(feature)) {
                return null;
            }

            return Entries.FirstOrDefault(entry => string.Equals(entry.Feature, feature.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Returns a Markdown capability matrix suitable for diagnostics and issue reports.</summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# OfficeIMO.Excel Format Capability Matrix");
            builder.AppendLine();
            builder.AppendLine("| Feature | XLSX | XLS | XLSB | Notes |");
            builder.AppendLine("| --- | --- | --- | --- | --- |");
            foreach (ExcelFormatCapabilityEntry entry in Entries) {
                builder.Append("| ");
                builder.Append(Escape(entry.Feature));
                builder.Append(" | ");
                builder.Append(entry.Xlsx);
                builder.Append(" | ");
                builder.Append(entry.Xls);
                builder.Append(" | ");
                builder.Append(entry.Xlsb);
                builder.Append(" | ");
                builder.Append(Escape(entry.Notes));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string Escape(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
