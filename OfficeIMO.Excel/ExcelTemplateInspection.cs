using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Template marker inspection result for a workbook or worksheet.
    /// </summary>
    public sealed class ExcelTemplateInspection {
        internal ExcelTemplateInspection(IReadOnlyList<ExcelTemplateMarkerInfo> markers, bool hasBindingInfo) {
            Markers = markers;
            HasBindingInfo = hasBindingInfo;
            TotalMarkers = markers.Count;
            UniqueMarkers = markers
                .Select(marker => marker.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            MissingMarkers = hasBindingInfo
                ? markers.Where(marker => marker.IsBound == false).ToList()
                : Array.Empty<ExcelTemplateMarkerInfo>();
            MissingMarkerNames = MissingMarkers
                .Select(marker => marker.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>Template markers discovered in workbook order.</summary>
        public IReadOnlyList<ExcelTemplateMarkerInfo> Markers { get; }

        /// <summary>True when the inspection was performed with a values/model binding source.</summary>
        public bool HasBindingInfo { get; }

        /// <summary>Total marker occurrences.</summary>
        public int TotalMarkers { get; }

        /// <summary>Distinct marker names discovered in the template.</summary>
        public IReadOnlyList<string> UniqueMarkers { get; }

        /// <summary>Marker occurrences missing from the supplied values/model. Empty when binding info was not supplied.</summary>
        public IReadOnlyList<ExcelTemplateMarkerInfo> MissingMarkers { get; }

        /// <summary>Distinct missing marker names. Empty when binding info was not supplied.</summary>
        public IReadOnlyList<string> MissingMarkerNames { get; }

        /// <summary>True when every marker was supplied by the inspected values/model.</summary>
        public bool AllMarkersBound => HasBindingInfo && MissingMarkers.Count == 0;

        /// <summary>
        /// Throws when this inspection was created without bindings or when any marker is missing.
        /// </summary>
        public ExcelTemplateInspection EnsureAllMarkersBound() {
            if (!HasBindingInfo) {
                throw new InvalidOperationException("Template inspection does not include binding information.");
            }

            if (MissingMarkerNames.Count > 0) {
                throw new InvalidOperationException("Missing template markers: " + string.Join(", ", MissingMarkerNames));
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of discovered template markers and binding status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Excel Template Markers");
            builder.AppendLine();
            builder.AppendLine($"Total markers: {TotalMarkers}");
            builder.AppendLine($"Unique markers: {UniqueMarkers.Count}");
            if (HasBindingInfo) {
                builder.AppendLine($"Missing markers: {MissingMarkerNames.Count}");
            }

            builder.AppendLine();
            builder.AppendLine("| Sheet | Cell | Marker | Format | Whole cell | Bound | Value kind |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (ExcelTemplateMarkerInfo marker in Markers) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(marker.SheetName));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(marker.CellReference));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(marker.Name));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(marker.Format ?? string.Empty));
                builder.Append(" | ");
                builder.Append(marker.IsWholeCell ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(marker.IsBound.HasValue ? marker.IsBound.Value ? "yes" : "no" : string.Empty);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(marker.BoundValueKind ?? string.Empty));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }

    /// <summary>
    /// Template marker metadata for a single marker occurrence.
    /// </summary>
    public sealed class ExcelTemplateMarkerInfo {
        internal ExcelTemplateMarkerInfo(string sheetName, string cellReference, string name, string? format, string cellText, bool isWholeCell,
            bool? isBound, string? boundValueKind, string? boundValueTypeName) {
            SheetName = sheetName;
            CellReference = cellReference;
            Name = name;
            Format = format;
            CellText = cellText;
            IsWholeCell = isWholeCell;
            IsBound = isBound;
            BoundValueKind = boundValueKind;
            BoundValueTypeName = boundValueTypeName;
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 cell reference.</summary>
        public string CellReference { get; }

        /// <summary>Marker name without braces or format suffix.</summary>
        public string Name { get; }

        /// <summary>Optional marker format, for example currency, percent, or yyyy-MM-dd.</summary>
        public string? Format { get; }

        /// <summary>Full cell text containing the marker.</summary>
        public string CellText { get; }

        /// <summary>True when the marker is the only content in the cell.</summary>
        public bool IsWholeCell { get; }

        /// <summary>True/false when inspected with bindings; null when no bindings were supplied.</summary>
        public bool? IsBound { get; }

        /// <summary>Coarse bound value category such as text, number, date/time, image, or null.</summary>
        public string? BoundValueKind { get; }

        /// <summary>CLR type name for the bound value. Null when the marker is unbound or inspection had no binding source.</summary>
        public string? BoundValueTypeName { get; }
    }
}
