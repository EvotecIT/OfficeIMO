using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        private static readonly Lazy<byte[]> ChartStyle251Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-style-251.xml"));

        private static readonly Lazy<byte[]> ChartColorStyle10Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-colors-10.xml"));

        private static readonly Lazy<byte[]> ChartTemplateBarBytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-template-bar.xml"));

        private static readonly Lazy<byte[]> ChartWorkbookBarBytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-workbook-bar.xlsx"));

        internal static void PopulateChartStyle(ChartStylePart stylePart) {
            if (stylePart == null) {
                throw new ArgumentNullException(nameof(stylePart));
            }

            using var stream = new MemoryStream(ChartStyle251Bytes.Value);
            stylePart.FeedData(stream);
        }

        internal static void PopulateChartColorStyle(ChartColorStylePart colorStylePart) {
            if (colorStylePart == null) {
                throw new ArgumentNullException(nameof(colorStylePart));
            }

            using var stream = new MemoryStream(ChartColorStyle10Bytes.Value);
            colorStylePart.FeedData(stream);
        }

        internal static void PopulateChartTemplate(ChartPart chartPart, string embeddedRelId, PowerPointChartData? data = null) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }

            using var stream = new MemoryStream(ChartTemplateBarBytes.Value);
            XDocument chartDoc = XDocument.Load(stream);
            XNamespace chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            var axisElements = chartDoc
                .Descendants()
                .Where(e => e.Name == chartNs + "axId" || e.Name == chartNs + "crossAx")
                .ToList();

            if (axisElements.Count > 0) {
                var axisMap = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var axisElement in axisElements) {
                    XAttribute? valAttribute = axisElement.Attribute("val");
                    if (valAttribute == null || string.IsNullOrWhiteSpace(valAttribute.Value)) {
                        continue;
                    }

                    if (!axisMap.TryGetValue(valAttribute.Value, out string? mapped)) {
                        mapped = PowerPointChartAxisIdGenerator.GetNextId().ToString(CultureInfo.InvariantCulture);
                        axisMap[valAttribute.Value] = mapped;
                    }

                    valAttribute.Value = mapped;
                }
            }

            // PowerPoint expects unique series identifiers per chart instance.
            XNamespace c16Ns = "http://schemas.microsoft.com/office/drawing/2014/chart";
            foreach (var uniqueId in chartDoc.Descendants(c16Ns + "uniqueId")) {
                uniqueId.SetAttributeValue("val", Guid.NewGuid().ToString("B").ToUpperInvariant());
            }

            if (!string.IsNullOrWhiteSpace(embeddedRelId)) {
                var externalData = chartDoc.Descendants(chartNs + "externalData").FirstOrDefault();
                if (externalData != null) {
                    externalData.SetAttributeValue(XName.Get("id", relNs.NamespaceName), embeddedRelId);
                }
            }

            if (data != null) {
                UpdateChartSeries(chartDoc, chartNs, data);
            }

            using var output = new MemoryStream();
            chartDoc.Save(output);
            output.Position = 0;
            chartPart.FeedData(output);
        }

        internal static byte[] BuildChartWorkbook(PowerPointChartData data) {
            byte[] template = GetChartWorkbookTemplateBytes();
            using MemoryStream ms = new();
            ms.Write(template, 0, template.Length);
            ms.Position = 0;

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, true)) {
                WorkbookPart wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("Chart workbook missing workbook part.");
                WorksheetPart wsPart = wbPart.WorksheetParts.FirstOrDefault()
                    ?? throw new InvalidOperationException("Chart workbook missing worksheet part.");

                var sheetData = wsPart.Worksheet.GetFirstChild<S.SheetData>() ?? new S.SheetData();
                sheetData.RemoveAllChildren<S.Row>();

                var sharedStringsPart = wbPart.SharedStringTablePart ?? wbPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable ??= new S.SharedStringTable();
                sharedStringsPart.SharedStringTable.RemoveAllChildren<S.SharedStringItem>();

                var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndex.TryGetValue(value, out int idx)) {
                        idx = stringIndex.Count;
                        stringIndex[value] = idx;
                        sharedStringsPart.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(value)));
                    }
                    return idx;
                }

                int seriesCount = data.Series.Count;
                int categoryCount = data.Categories.Count;
                int totalColumns = seriesCount + 1;
                int totalRows = categoryCount + 1;
                string lastColumn = ColumnLetter(totalColumns);
                string dimensionRef = $"A1:{lastColumn}{totalRows}";

                var headerRow = new S.Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                headerRow.Append(CreateSharedStringCell("A1", GetStringIndex(" ")));
                for (int i = 0; i < seriesCount; i++) {
                    string cellRef = $"{ColumnLetter(i + 2)}1";
                    headerRow.Append(CreateSharedStringCell(cellRef, GetStringIndex(data.Series[i].Name)));
                }
                sheetData.Append(headerRow);

                for (int rowIndex = 0; rowIndex < categoryCount; rowIndex++) {
                    uint excelRow = (uint)(rowIndex + 2);
                    var row = new S.Row { RowIndex = excelRow, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                    string category = data.Categories[rowIndex] ?? string.Empty;
                    row.Append(CreateSharedStringCell($"A{excelRow}", GetStringIndex(category)));

                    for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        string cellRef = $"{ColumnLetter(seriesIndex + 2)}{excelRow}";
                        double value = data.Series[seriesIndex].Values[rowIndex];
                        row.Append(CreateNumberCell(cellRef, value));
                    }

                    sheetData.Append(row);
                }

                if (wsPart.Worksheet.GetFirstChild<S.SheetData>() == null) {
                    wsPart.Worksheet.Append(sheetData);
                }

                var dimension = wsPart.Worksheet.SheetDimension ?? wsPart.Worksheet.GetFirstChild<S.SheetDimension>();
                if (dimension == null) {
                    dimension = new S.SheetDimension();
                    wsPart.Worksheet.InsertAt(dimension, 0);
                }
                dimension.Reference = dimensionRef;

                var cols = wsPart.Worksheet.GetFirstChild<S.Columns>();
                if (cols != null) {
                    var colList = cols.Elements<S.Column>().ToList();
                    if (colList.Count > 0) {
                        colList[0].Min = 1U;
                        colList[0].Max = 1U;
                    }
                    if (colList.Count > 1) {
                        colList[1].Min = 2U;
                        colList[1].Max = (uint)totalColumns;
                    } else if (totalColumns > 1) {
                        cols.Append(new S.Column { Min = 2U, Max = (uint)totalColumns, Width = 9.36328125, CustomWidth = true });
                    }
                }

                var tablePart = wsPart.TableDefinitionParts.FirstOrDefault();
                if (tablePart?.Table != null) {
                    tablePart.Table.Reference = dimensionRef;
                    var tableColumns = tablePart.Table.TableColumns ?? new S.TableColumns();
                    tableColumns.RemoveAllChildren<S.TableColumn>();
                    tableColumns.Append(new S.TableColumn { Id = 1U, Name = " " });
                    for (int i = 0; i < seriesCount; i++) {
                        tableColumns.Append(new S.TableColumn { Id = (uint)(i + 2), Name = data.Series[i].Name });
                    }
                    tableColumns.Count = (uint)totalColumns;
                    tablePart.Table.TableColumns = tableColumns;
                    tablePart.Table.Save();
                }

                sharedStringsPart.SharedStringTable.Count = (uint)stringIndex.Count;
                sharedStringsPart.SharedStringTable.UniqueCount = (uint)stringIndex.Count;
                sharedStringsPart.SharedStringTable.Save();
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        internal static byte[] GetChartWorkbookTemplateBytes() {
            return ChartWorkbookBarBytes.Value;
        }

        private static void UpdateChartSeries(XDocument chartDoc, XNamespace chartNs, PowerPointChartData data) {
            var barChart = chartDoc.Descendants(chartNs + "barChart").FirstOrDefault();
            if (barChart == null) {
                return;
            }

            var seriesElements = barChart.Elements(chartNs + "ser").ToList();
            if (seriesElements.Count == 0) {
                return;
            }

            var insertBefore = barChart.Elements().FirstOrDefault(e => e.Name != chartNs + "ser");
            var newSeries = new List<XElement>();
            for (int i = 0; i < data.Series.Count; i++) {
                var template = i < seriesElements.Count ? seriesElements[i] : seriesElements.Last();
                var seriesElement = new XElement(template);
                UpdateSeriesElement(seriesElement, chartNs, data.Series[i], data.Categories, i);
                newSeries.Add(seriesElement);
            }

            barChart.Elements(chartNs + "ser").Remove();
            foreach (var seriesElement in newSeries) {
                if (insertBefore != null) {
                    insertBefore.AddBeforeSelf(seriesElement);
                } else {
                    barChart.Add(seriesElement);
                }
            }
        }

        private static void UpdateSeriesElement(XElement seriesElement, XNamespace chartNs, PowerPointChartSeries series,
            IReadOnlyList<string> categories, int seriesIndex) {
            int lastRow = categories.Count + 1;
            string seriesCol = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesCol}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesCol}$2:${seriesCol}${lastRow}";

            SetValue(seriesElement, chartNs, "order", seriesIndex);
            SetValue(seriesElement, chartNs, "idx", seriesIndex);

            var tx = seriesElement.Element(chartNs + "tx") ?? new XElement(chartNs + "tx");
            var txRef = tx.Element(chartNs + "strRef") ?? new XElement(chartNs + "strRef");
            SetStringRef(chartNs, txRef, seriesNameRef, new[] { series.Name });
            if (tx.Parent == null) tx.Add(txRef);
            if (tx.Parent == null) seriesElement.Add(tx);
            else if (txRef.Parent == null) tx.Add(txRef);

            var cat = seriesElement.Element(chartNs + "cat") ?? new XElement(chartNs + "cat");
            var catRef = cat.Element(chartNs + "strRef") ?? new XElement(chartNs + "strRef");
            SetStringRef(chartNs, catRef, categoriesRef, categories);
            if (cat.Parent == null) cat.Add(catRef);
            if (cat.Parent == null) seriesElement.Add(cat);
            else if (catRef.Parent == null) cat.Add(catRef);

            var val = seriesElement.Element(chartNs + "val") ?? new XElement(chartNs + "val");
            var numRef = val.Element(chartNs + "numRef") ?? new XElement(chartNs + "numRef");
            SetNumberRef(chartNs, numRef, valuesRef, series.Values);
            if (val.Parent == null) val.Add(numRef);
            if (val.Parent == null) seriesElement.Add(val);
            else if (numRef.Parent == null) val.Add(numRef);

            XNamespace c16 = "http://schemas.microsoft.com/office/drawing/2014/chart";
            var uniqueId = seriesElement.Descendants(c16 + "uniqueId").FirstOrDefault();
            if (uniqueId != null) {
                uniqueId.SetAttributeValue("val", Guid.NewGuid().ToString("B").ToUpperInvariant());
            }
        }

        private static void SetValue(XElement seriesElement, XNamespace chartNs, string name, int value) {
            var element = seriesElement.Element(chartNs + name);
            if (element == null) {
                element = new XElement(chartNs + name);
                seriesElement.AddFirst(element);
            }
            element.SetAttributeValue("val", value.ToString(CultureInfo.InvariantCulture));
        }

        private static void SetStringRef(XNamespace chartNs, XElement strRef, string formula, IReadOnlyList<string> values) {
            var f = strRef.Element(chartNs + "f");
            if (f == null) {
                f = new XElement(chartNs + "f");
                strRef.AddFirst(f);
            }
            f.Value = formula;

            var cache = strRef.Element(chartNs + "strCache");
            if (cache == null) {
                cache = new XElement(chartNs + "strCache");
                strRef.Add(cache);
            }

            cache.RemoveNodes();
            cache.Add(new XElement(chartNs + "ptCount", new XAttribute("val", values.Count)));
            for (int i = 0; i < values.Count; i++) {
                cache.Add(new XElement(chartNs + "pt",
                    new XAttribute("idx", i.ToString(CultureInfo.InvariantCulture)),
                    new XElement(chartNs + "v", values[i] ?? string.Empty)));
            }
        }

        private static void SetNumberRef(XNamespace chartNs, XElement numRef, string formula, IReadOnlyList<double> values) {
            var f = numRef.Element(chartNs + "f");
            if (f == null) {
                f = new XElement(chartNs + "f");
                numRef.AddFirst(f);
            }
            f.Value = formula;

            var cache = numRef.Element(chartNs + "numCache");
            string formatCode = cache?.Element(chartNs + "formatCode")?.Value ?? "General";
            if (cache == null) {
                cache = new XElement(chartNs + "numCache");
                numRef.Add(cache);
            }

            cache.RemoveNodes();
            cache.Add(new XElement(chartNs + "formatCode", formatCode));
            cache.Add(new XElement(chartNs + "ptCount", new XAttribute("val", values.Count)));
            for (int i = 0; i < values.Count; i++) {
                cache.Add(new XElement(chartNs + "pt",
                    new XAttribute("idx", i.ToString(CultureInfo.InvariantCulture)),
                    new XElement(chartNs + "v", values[i].ToString(CultureInfo.InvariantCulture))));
            }
        }

        private static string ColumnLetter(int column) {
            int dividend = column;
            string columnName = string.Empty;
            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static S.Cell CreateSharedStringCell(string cellReference, int sharedStringIndex) {
            return new S.Cell {
                CellReference = cellReference,
                DataType = S.CellValues.SharedString,
                CellValue = new S.CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture))
            };
        }

        private static S.Cell CreateNumberCell(string cellReference, double value) {
            return new S.Cell {
                CellReference = cellReference,
                CellValue = new S.CellValue(value.ToString(CultureInfo.InvariantCulture))
            };
        }

    }
}
