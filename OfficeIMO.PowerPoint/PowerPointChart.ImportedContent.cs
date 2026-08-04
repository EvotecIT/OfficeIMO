using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

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
            PowerPointImportedChartSupport support, OfficeChartKind? projection,
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
        public OfficeChartKind? ExportProjection { get; }
        /// <summary>Actionable preservation or fidelity diagnostics.</summary>
        public IReadOnlyList<string> Diagnostics { get; }
    }

    public partial class PowerPointChart {
        private static readonly IReadOnlyDictionary<string, OfficeChartKind>
            AdvancedChartProjections = new Dictionary<string, OfficeChartKind>(StringComparer.Ordinal) {
                ["bar3DChart"] = OfficeChartKind.ColumnClustered,
                ["line3DChart"] = OfficeChartKind.Line,
                ["area3DChart"] = OfficeChartKind.Area,
                ["pie3DChart"] = OfficeChartKind.Pie,
                ["ofPieChart"] = OfficeChartKind.Pie,
                ["stockChart"] = OfficeChartKind.Line,
                ["surfaceChart"] = OfficeChartKind.Line,
                ["surface3DChart"] = OfficeChartKind.Line
            };

        /// <summary>Inspects native family, typed-edit support, and export fidelity.</summary>
        public PowerPointImportedChartReport InspectImportedContent() {
            C.PlotArea? plotArea = GetChart().GetFirstChild<C.PlotArea>();
            OpenXmlElement[] groups = plotArea?.ChildElements
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal))
                .ToArray() ?? Array.Empty<OpenXmlElement>();
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
                OfficeChartKind projection = GetAdvancedProjection(groups[0]);
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

        private string? DescribeAdvancedEditBlocker(OpenXmlElement group) {
            if (!TryGetEditableImportedWorkbook(out _, out string workbookDiagnostic))
                return $"Imported {group.LocalName} {workbookDiagnostic}";
            OpenXmlCompositeElement[] series = group.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName == "ser").ToArray();
            if (series.Length == 0)
                return $"Imported {group.LocalName} has no safely editable series and remains preservation-only.";
            bool unsupported = series.Any(item =>
                item.GetFirstChild<C.SeriesText>()?.GetFirstChild<C.StringReference>() == null ||
                item.GetFirstChild<C.CategoryAxisData>()?.GetFirstChild<C.StringReference>() == null ||
                item.GetFirstChild<C.Values>()?.GetFirstChild<C.NumberReference>() == null);
            return unsupported
                ? $"Imported {group.LocalName} uses producer-specific literal or numeric category storage; editing is rejected so the native data model remains unchanged."
                : null;
        }

        internal static OfficeChartKind GetAdvancedProjection(OpenXmlElement group) {
            if (group is C.Bar3DChart bar) {
                bool horizontal = bar.GetFirstChild<C.BarDirection>()?.Val?.Value == C.BarDirectionValues.Bar;
                C.BarGroupingValues grouping = bar.GetFirstChild<C.BarGrouping>()?.Val?.Value ??
                    C.BarGroupingValues.Clustered;
                if (grouping == C.BarGroupingValues.Stacked)
                    return horizontal ? OfficeChartKind.BarStacked : OfficeChartKind.ColumnStacked;
                if (grouping == C.BarGroupingValues.PercentStacked)
                    return horizontal ? OfficeChartKind.BarStacked100 : OfficeChartKind.ColumnStacked100;
                return horizontal ? OfficeChartKind.BarClustered : OfficeChartKind.ColumnClustered;
            }
            if (group is C.Line3DChart line) {
                C.GroupingValues grouping = line.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
                return grouping == C.GroupingValues.Stacked ? OfficeChartKind.LineStacked :
                    grouping == C.GroupingValues.PercentStacked ? OfficeChartKind.LineStacked100 : OfficeChartKind.Line;
            }
            if (group is C.Area3DChart area) {
                C.GroupingValues grouping = area.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
                return grouping == C.GroupingValues.Stacked ? OfficeChartKind.AreaStacked :
                    grouping == C.GroupingValues.PercentStacked ? OfficeChartKind.AreaStacked100 : OfficeChartKind.Area;
            }
            return AdvancedChartProjections[group.LocalName];
        }

        /// <summary>
        /// Updates series names, categories, values, formulas, caches, and the embedded
        /// workbook of an advanced imported chart without rebuilding its native chart group.
        /// </summary>
        public PowerPointChart UpdateImportedData(OfficeChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            PowerPointImportedChartReport report = InspectImportedContent();
            if (report.Support != PowerPointImportedChartSupport.EditableWithProjectedRendering ||
                report.ExportProjection == null)
                throw new NotSupportedException(report.Diagnostics.FirstOrDefault() ??
                    "This imported chart does not require the advanced in-place editor.");
            PowerPointUtils.ValidateSharedChartData(data, report.ExportProjection.Value);
            ChartPart chartPart = GetChartPart();
            if (!TryGetEditableImportedWorkbook(out EmbeddedPackagePart? embedded,
                    out string workbookDiagnostic) || embedded == null)
                throw new NotSupportedException(workbookDiagnostic);
            C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            OpenXmlCompositeElement group = (OpenXmlCompositeElement)plotArea.ChildElements
                .Single(element => element.LocalName == report.Family);
            List<OpenXmlCompositeElement> series = group.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName == "ser").ToList();
            if (series.Count != data.Series.Count)
                throw new NotSupportedException("Changing the series count of an imported advanced chart can alter its meaning; update the existing series only.");
            byte[] original;
            using (Stream stream = embedded.GetStream(FileMode.Open,
                       FileAccess.Read)) {
                original = PowerPointChartWorkbookSecurity.ReadAndValidate(
                    stream);
            }
            byte[] workbook = PowerPointChartWorkbookEditor.Update(original,
                data);
            using (var validation = new MemoryStream(workbook,
                       writable: false)) {
                _ = PowerPointChartWorkbookSecurity.ReadAndValidate(
                    validation);
            }
            C.ChartSpace originalChartSpace = (C.ChartSpace)chartPart
                .ChartSpace.CloneNode(true);
            try {
                int lastRow = data.Categories.Count + 1;
                for (int index = 0; index < series.Count; index++) {
                    OfficeChartSeries item = data.Series[index];
                    string column = GetExcelColumn(index + 2);
                    UpdateStringReference(series[index]
                            .GetFirstChild<C.SeriesText>()!,
                        $"Sheet1!${column}$1", new[] { item.Name });
                    UpdateStringReference(series[index]
                            .GetFirstChild<C.CategoryAxisData>()!,
                        $"Sheet1!$A$2:$A${lastRow}", data.Categories);
                    UpdateNumberReference(series[index]
                            .GetFirstChild<C.Values>()!,
                        $"Sheet1!${column}$2:${column}${lastRow}",
                        item.Values);
                }
                using var replacement = new MemoryStream(workbook,
                    writable: false);
                embedded.FeedData(replacement);
                Save();
            } catch {
                try {
                    chartPart.ChartSpace = originalChartSpace;
                    chartPart.ChartSpace.Save();
                } catch {
                    // Best-effort restoration continues with the workbook.
                }
                try {
                    using var rollback = new MemoryStream(original,
                        writable: false);
                    embedded.FeedData(rollback);
                } catch {
                    // Preserve the original update exception. A subsequent
                    // save/validation will surface any package I/O failure.
                }
                throw;
            }
            return this;
        }

        private bool TryGetEditableImportedWorkbook(
            out EmbeddedPackagePart? workbook, out string diagnostic) {
            ChartPart chartPart = GetChartPart();
            C.ExternalData[] references = chartPart.ChartSpace?
                .Descendants<C.ExternalData>().ToArray() ?? Array.Empty<C.ExternalData>();
            if (references.Length != 1 ||
                string.IsNullOrWhiteSpace(references[0].Id?.Value) ||
                !chartPart.TryGetPartById(references[0].Id!.Value!,
                    out OpenXmlPart? referencedPart) ||
                referencedPart is not EmbeddedPackagePart embedded) {
                workbook = null;
                diagnostic = "has no single referenced embedded workbook that can be updated consistently; package markup is preserved unchanged.";
                return false;
            }
            if (!PowerPointPresentation.IsSafeChartWorkbookPart(chartPart,
                    embedded)) {
                workbook = null;
                diagnostic = "uses a richer or producer-specific workbook; editing is rejected so sheets, formulas, formatting, macros, and related package content remain unchanged.";
                return false;
            }
            workbook = embedded;
            diagnostic = string.Empty;
            return true;
        }

        private static void UpdateStringReference(OpenXmlCompositeElement holder,
            string formula, IReadOnlyList<string> values) {
            if (holder == null) throw new NotSupportedException("The advanced chart series does not use string-backed category data.");
            C.StringReference reference = holder.GetFirstChild<C.StringReference>() ??
                throw new NotSupportedException("The advanced chart series uses producer-specific string storage.");
            reference.GetFirstChild<C.Formula>()?.Remove();
            reference.InsertAt(new C.Formula(formula), 0);
            reference.GetFirstChild<C.StringCache>()?.Remove();
            var cache = new C.StringCache(new C.PointCount { Val = (uint)values.Count });
            for (int index = 0; index < values.Count; index++)
                cache.Append(new C.StringPoint(new C.NumericValue(values[index] ?? string.Empty)) { Index = (uint)index });
            reference.Append(cache);
        }

        private static void UpdateNumberReference(OpenXmlCompositeElement holder,
            string formula, IReadOnlyList<double> values) {
            if (holder == null) throw new NotSupportedException("The advanced chart series does not expose numeric values.");
            C.NumberReference reference = holder.GetFirstChild<C.NumberReference>() ??
                throw new NotSupportedException("The advanced chart series uses producer-specific numeric storage.");
            string formatCode = reference.GetFirstChild<C.NumberingCache>()?
                .GetFirstChild<C.FormatCode>()?.Text ?? "General";
            reference.GetFirstChild<C.Formula>()?.Remove();
            reference.InsertAt(new C.Formula(formula), 0);
            reference.GetFirstChild<C.NumberingCache>()?.Remove();
            var cache = new C.NumberingCache(new C.FormatCode(formatCode),
                new C.PointCount { Val = (uint)values.Count });
            for (int index = 0; index < values.Count; index++)
                cache.Append(new C.NumericPoint(new C.NumericValue(values[index].ToString("R", CultureInfo.InvariantCulture))) { Index = (uint)index });
            reference.Append(cache);
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
