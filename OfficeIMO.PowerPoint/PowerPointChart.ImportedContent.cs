using System.Collections.ObjectModel;
using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>Support level for one imported native chart family.</summary>
    public enum PowerPointImportedChartSupport {
        /// <summary>Data can be edited in place and export has an exact semantic renderer.</summary>
        EditableAndRenderable,
        /// <summary>Data can be edited in place; export uses a documented semantic projection.</summary>
        EditableWithProjectedRendering,
        /// <summary>The package content is retained but is not safe to edit or render.</summary>
        PreservationOnly
    }

    /// <summary>Inspection result for imported chart content.</summary>
    public sealed class PowerPointImportedChartReport {
        internal PowerPointImportedChartReport(string family,
            PowerPointImportedChartSupport support, OfficeChartType? projection,
            IEnumerable<string> diagnostics) {
            Family = family;
            Support = support;
            ExportProjection = projection;
            Diagnostics = new ReadOnlyCollection<string>(diagnostics.ToList());
        }
        /// <summary>Native Open XML chart family.</summary>
        public string Family { get; }
        /// <summary>Typed edit/export support.</summary>
        public PowerPointImportedChartSupport Support { get; }
        /// <summary>Semantic image/PDF projection when exact rendering is unavailable.</summary>
        public OfficeChartType? ExportProjection { get; }
        /// <summary>Actionable preservation or fidelity diagnostics.</summary>
        public IReadOnlyList<string> Diagnostics { get; }
    }

    public partial class PowerPointChart {
        private static readonly IReadOnlyDictionary<string, OfficeChartType>
            AdvancedChartProjections = new Dictionary<string, OfficeChartType>(StringComparer.Ordinal) {
                ["bar3DChart"] = OfficeChartType.ColumnClustered,
                ["line3DChart"] = OfficeChartType.Line,
                ["area3DChart"] = OfficeChartType.Area,
                ["pie3DChart"] = OfficeChartType.Pie,
                ["ofPieChart"] = OfficeChartType.Pie,
                ["stockChart"] = OfficeChartType.Line,
                ["surfaceChart"] = OfficeChartType.Line,
                ["surface3DChart"] = OfficeChartType.Line
            };

        /// <summary>Inspects native family, typed-edit support, and export fidelity.</summary>
        public PowerPointImportedChartReport InspectImportedContent() {
            OfficeXmlNode plotArea = ChartNode.Child(CNs, "plotArea");
            OfficeXmlNode[] groups = plotArea.Children().ToList()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal))
                .ToArray();
            if (groups.Length != 1) {
                bool renderableMixed = groups.Length > 1 &&
                    TryGetOfficeSnapshot(out _);
                return new PowerPointImportedChartReport(groups.Length == 0 ? "none" : "mixed",
                    renderableMixed ? PowerPointImportedChartSupport.EditableAndRenderable :
                    PowerPointImportedChartSupport.PreservationOnly, null,
                    groups.Length == 0
                        ? new[] { "No native chart group was found." }
                        : renderableMixed
                            ? Array.Empty<string>()
                            : new[] { "The imported mixed-chart combination cannot be projected without changing its meaning and remains preservation-only." });
            }
            string family = groups[0].LocalName;
            if (AdvancedChartProjections.ContainsKey(family)) {
                OfficeChartType projection = GetAdvancedProjection(groups[0]);
                string? editBlocker = DescribeAdvancedEditBlocker(groups[0]);
                if (editBlocker != null) {
                    return new PowerPointImportedChartReport(family,
                        PowerPointImportedChartSupport.PreservationOnly, projection,
                        new[] { editBlocker });
                }
                return new PowerPointImportedChartReport(family,
                    PowerPointImportedChartSupport.EditableWithProjectedRendering,
                    projection, new[] {
                        $"Image and PDF export use the {projection} semantic projection; native {family} geometry remains preserved for PowerPoint."
                    });
            }
            if (family == "barChart" || family == "lineChart" || family == "areaChart" ||
                family == "radarChart" || family == "scatterChart" || family == "bubbleChart" ||
                family == "pieChart" || family == "doughnutChart") {
                return new PowerPointImportedChartReport(family,
                    PowerPointImportedChartSupport.EditableAndRenderable, null,
                    Array.Empty<string>());
            }
            return new PowerPointImportedChartReport(family,
                PowerPointImportedChartSupport.PreservationOnly, null,
                new[] { $"Producer-specific chart family '{family}' is preserved losslessly but is not modeled for editing or export." });
        }

        /// <summary>The <c>ser</c> children of a chart group, in any namespace, as the imported-chart checks read them.</summary>
        private static List<OfficeXmlNode> GroupSeries(OfficeXmlNode group) =>
            group.Children().ToList().Where(element => element.LocalName == "ser").ToList();

        private string? DescribeAdvancedEditBlocker(OfficeXmlNode group) {
            if (!TryGetEditableImportedWorkbook(out _, out string workbookDiagnostic))
                return $"Imported {group.LocalName} {workbookDiagnostic}";
            List<OfficeXmlNode> series = GroupSeries(group);
            if (series.Count == 0)
                return $"Imported {group.LocalName} has no safely editable series and remains preservation-only.";
            bool unsupported = series.Any(item =>
                item.Child(CNs, "tx").Child(CNs, "strRef").IsNone ||
                item.Child(CNs, "cat").Child(CNs, "strRef").IsNone ||
                item.Child(CNs, "val").Child(CNs, "numRef").IsNone);
            return unsupported
                ? $"Imported {group.LocalName} uses producer-specific literal or numeric category storage; editing is rejected so the native data model remains unchanged."
                : null;
        }

        internal static OfficeChartType GetAdvancedProjection(OfficeXmlNode group) {
            if (group.Is(CNs, "bar3DChart")) {
                bool horizontal = group.ChildVal(CNs, "barDir") == "bar";
                string? grouping = group.ChildVal(CNs, "grouping");
                if (grouping == "stacked")
                    return horizontal ? OfficeChartType.BarStacked : OfficeChartType.ColumnStacked;
                if (grouping == "percentStacked")
                    return horizontal ? OfficeChartType.BarStacked100 : OfficeChartType.ColumnStacked100;
                return horizontal ? OfficeChartType.BarClustered : OfficeChartType.ColumnClustered;
            }
            if (group.Is(CNs, "line3DChart")) {
                string? grouping = group.ChildVal(CNs, "grouping");
                return grouping == "stacked" ? OfficeChartType.LineStacked :
                    grouping == "percentStacked" ? OfficeChartType.LineStacked100 : OfficeChartType.Line;
            }
            if (group.Is(CNs, "area3DChart")) {
                string? grouping = group.ChildVal(CNs, "grouping");
                return grouping == "stacked" ? OfficeChartType.AreaStacked :
                    grouping == "percentStacked" ? OfficeChartType.AreaStacked100 : OfficeChartType.Area;
            }
            return AdvancedChartProjections[group.LocalName];
        }

        /// <summary>
        /// Updates series names, categories, values, formulas, caches, and the embedded
        /// workbook of an advanced imported chart without rebuilding its native chart group.
        /// </summary>
        public PowerPointChart UpdateImportedData(OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(data);
            PowerPointImportedChartReport report = InspectImportedContent();
            if (report.Support != PowerPointImportedChartSupport.EditableWithProjectedRendering ||
                report.ExportProjection == null)
                throw new NotSupportedException(report.Diagnostics.FirstOrDefault() ??
                    "This imported chart does not require the advanced in-place editor.");
            PowerPointUtils.ValidateSharedChartData(data, report.ExportProjection.Value);
            IOfficeChartPart chartPart = SharedChartPart;
            if (!TryGetEditableImportedWorkbook(out IOfficeChartEmbeddedPackage? embedded,
                    out string workbookDiagnostic) || embedded == null)
                throw new NotSupportedException(workbookDiagnostic);
            OfficeXmlNode chartSpace = chartPart.ChartSpace;
            OfficeXmlNode plotArea = chartSpace.Child(CNs, "chart").Child(CNs, "plotArea");
            OfficeXmlNode group = plotArea.Children().ToList()
                .Single(element => element.LocalName == report.Family);
            List<OfficeXmlNode> series = GroupSeries(group);
            if (series.Count != data.Series.Count)
                throw new NotSupportedException("Changing the series count of an imported advanced chart can alter its meaning; update the existing series only.");
            byte[] original;
            using (Stream stream = embedded.OpenRead()) {
                original = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(stream);
            }
            byte[] workbook = PowerPointChartWorkbookEditor.Update(original,
                data);
            Action preserveBindings = OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                chartPart, data.Series.Select(item => item.Values.Count).ToArray());
            using (var validation = new MemoryStream(workbook,
                       writable: false)) {
                _ = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(
                    validation);
            }
            OfficeXmlNode originalChartSpace = chartSpace.Clone();
            try {
                int lastRow = data.Categories.Count + 1;
                for (int index = 0; index < series.Count; index++) {
                    OfficeSeries item = data.Series[index];
                    string column = GetExcelColumn(index + 2);
                    UpdateStringReference(series[index].Child(CNs, "tx"),
                        $"Sheet1!${column}$1", new[] { item.Name });
                    UpdateStringReference(series[index].Child(CNs, "cat"),
                        $"Sheet1!$A$2:$A${lastRow}", data.Categories);
                    UpdateNumberReference(series[index].Child(CNs, "val"),
                        $"Sheet1!${column}$2:${column}${lastRow}",
                        item.Values);
                }
                preserveBindings();
                embedded.Replace(workbook);
                MarkChartEdited();
            } catch {
                try {
                    // The chart space goes back to its content before the update.
                    chartSpace.RemoveAllChildElements();
                    foreach (OfficeXmlNode child in originalChartSpace.Children().ToList()) {
                        child.Remove();
                        chartSpace.Append(child);
                    }
                    MarkChartEdited();
                } catch {
                    // Best-effort restoration continues with the workbook.
                }
                try {
                    embedded.Replace(original);
                } catch {
                    // Preserve the original update exception. A subsequent
                    // save/validation will surface any package I/O failure.
                }
                throw;
            }
            return this;
        }

        private bool TryGetEditableImportedWorkbook(
            out IOfficeChartEmbeddedPackage? workbook, out string diagnostic) {
            IOfficeChartPart chartPart = SharedChartPart;
            List<OfficeXmlNode> references = chartPart.ChartSpace.ContentDescendants(CNs, "externalData");
            string? id = references.Count == 1 ? references[0].Attribute(OfficeXmlNamespaces.Relationships, "id") : null;
            IOfficeChartEmbeddedPackage? embedded = string.IsNullOrWhiteSpace(id) ? null
                : chartPart.EmbeddedPackages(int.MaxValue).FirstOrDefault(package => string.Equals(package.RelationshipId, id, StringComparison.Ordinal));
            if (embedded == null) {
                workbook = null;
                diagnostic = "has no single referenced embedded workbook that can be updated consistently; package markup is preserved unchanged.";
                return false;
            }
            if (!PowerPointDocument.IsSafeChartWorkbookPart(chartPart, embedded)) {
                workbook = null;
                diagnostic = "uses a richer or producer-specific workbook; editing is rejected so sheets, formulas, formatting, macros, and related package content remain unchanged.";
                return false;
            }
            workbook = embedded;
            diagnostic = string.Empty;
            return true;
        }

        private static void UpdateStringReference(OfficeXmlNode holder,
            string formula, IReadOnlyList<string> values) {
            if (holder.IsNone) throw new NotSupportedException("The advanced chart series does not use string-backed category data.");
            OfficeXmlNode reference = holder.Child(CNs, "strRef");
            if (reference.IsNone) throw new NotSupportedException("The advanced chart series uses producer-specific string storage.");
            ReplaceFormula(reference, formula);
            OfficeXmlNode existing = reference.Child(CNs, "strCache");
            if (existing.Exists) existing.Remove();
            OfficeXmlNode cache = AppendNew(reference, CNs, "strCache");
            AppendNew(cache, CNs, "ptCount").SetUInt32Attribute("val", (uint)values.Count);
            for (int index = 0; index < values.Count; index++) {
                OfficeXmlNode point = AppendNew(cache, CNs, "pt");
                point.SetUInt32Attribute("idx", (uint)index);
                AppendNew(point, CNs, "v").SetInnerText(values[index] ?? string.Empty);
            }
        }

        private static void UpdateNumberReference(OfficeXmlNode holder,
            string formula, IReadOnlyList<double> values) {
            if (holder.IsNone) throw new NotSupportedException("The advanced chart series does not expose numeric values.");
            OfficeXmlNode reference = holder.Child(CNs, "numRef");
            if (reference.IsNone) throw new NotSupportedException("The advanced chart series uses producer-specific numeric storage.");
            string formatCode = reference.Child(CNs, "numCache").ChildText(CNs, "formatCode") ?? "General";
            ReplaceFormula(reference, formula);
            OfficeXmlNode existing = reference.Child(CNs, "numCache");
            if (existing.Exists) existing.Remove();
            OfficeXmlNode cache = AppendNew(reference, CNs, "numCache");
            AppendNew(cache, CNs, "formatCode").SetInnerText(formatCode);
            AppendNew(cache, CNs, "ptCount").SetUInt32Attribute("val", (uint)values.Count);
            for (int index = 0; index < values.Count; index++) {
                OfficeXmlNode point = AppendNew(cache, CNs, "pt");
                point.SetUInt32Attribute("idx", (uint)index);
                AppendNew(point, CNs, "v").SetInnerText(OfficeIMO.Core.Internal.OfficeRoundTripText.Format(values[index], CultureInfo.InvariantCulture));
            }
        }

        // Replaces the reference's first c:f with one holding the formula, written first.
        private static void ReplaceFormula(OfficeXmlNode reference, string formula) {
            OfficeXmlNode existing = reference.Child(CNs, "f");
            if (existing.Exists) existing.Remove();
            reference.InsertBefore(reference.CreateElement(CNs, "f"), reference.FirstChild).SetInnerText(formula);
        }

        private static string GetExcelColumn(int index) {
            string result = string.Empty;
            while (index > 0) {
                index--;
                result = (char)('A' + index % 26) + result;
                index /= 26;
            }
            return result;
        }
    }
}
