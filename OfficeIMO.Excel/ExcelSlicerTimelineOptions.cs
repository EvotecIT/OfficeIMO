using System.Security;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for authoring workbook-level slicer cache metadata.
    /// </summary>
    public sealed class ExcelSlicerCacheOptions {
        /// <summary>Slicer cache name.</summary>
        public string Name { get; set; } = "SlicerCache";

        /// <summary>Source field, table column, or pivot field name.</summary>
        public string? SourceName { get; set; }

        /// <summary>Optional pivot table name the slicer is intended to filter.</summary>
        public string? PivotTableName { get; set; }

        /// <summary>Optional caller-supplied XML. When set, OfficeIMO writes it as-is.</summary>
        public string? Xml { get; set; }

        internal string ToXml() {
            if (!string.IsNullOrWhiteSpace(Xml)) {
                return Xml!;
            }

            return "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"" +
                $" name=\"{Escape(Name)}\"" +
                OptionalAttribute("sourceName", SourceName) +
                OptionalAttribute("pivotTableName", PivotTableName) +
                "/>";
        }

        private static string OptionalAttribute(string name, string? value)
            => string.IsNullOrWhiteSpace(value) ? string.Empty : $" {name}=\"{Escape(value!)}\"";

        private static string Escape(string value) => SecurityElement.Escape(value) ?? string.Empty;
    }

    /// <summary>
    /// Options for authoring workbook-level timeline cache metadata.
    /// </summary>
    public sealed class ExcelTimelineCacheOptions {
        /// <summary>Timeline cache name.</summary>
        public string Name { get; set; } = "TimelineCache";

        /// <summary>Source date field, table column, or pivot field name.</summary>
        public string? SourceName { get; set; }

        /// <summary>Optional pivot table name the timeline is intended to filter.</summary>
        public string? PivotTableName { get; set; }

        /// <summary>Optional caller-supplied XML. When set, OfficeIMO writes it as-is.</summary>
        public string? Xml { get; set; }

        internal string ToXml() {
            if (!string.IsNullOrWhiteSpace(Xml)) {
                return Xml!;
            }

            return "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\"" +
                $" name=\"{Escape(Name)}\"" +
                OptionalAttribute("sourceName", SourceName) +
                OptionalAttribute("pivotTableName", PivotTableName) +
                "/>";
        }

        private static string OptionalAttribute(string name, string? value)
            => string.IsNullOrWhiteSpace(value) ? string.Empty : $" {name}=\"{Escape(value!)}\"";

        private static string Escape(string value) => SecurityElement.Escape(value) ?? string.Empty;
    }
}
