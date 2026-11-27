using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    /// <summary>Native ChartEx layouts authored and mutated by OfficeIMO.</summary>
    public enum ExcelModernChartType {
        /// <summary>Unsupported or imported layout not mapped by OfficeIMO.</summary>
        Unsupported = 0,
        /// <summary>Native funnel chart.</summary>
        Funnel,
        /// <summary>Native waterfall chart.</summary>
        Waterfall,
        /// <summary>Native box-and-whisker chart.</summary>
        BoxWhisker,
        /// <summary>Native treemap chart.</summary>
        Treemap,
        /// <summary>Native sunburst chart.</summary>
        Sunburst
    }

    /// <summary>Native or imported worksheet ChartEx drawing.</summary>
    public sealed class ExcelModernChart {
        private readonly OfficeXmlNode _frame;
        private readonly ExcelPartRef _drawingsPart;
        private readonly ExcelSheet _sheet;

        internal ExcelModernChart(
            OfficeXmlNode frame,
            ExcelPartRef drawingsPart,
            ExcelSheet sheet,
            ExcelChartDataRange? dataRange = null) {
            _frame = frame;
            _drawingsPart = drawingsPart;
            _sheet = sheet;
            DataRange = dataRange;
        }

        private OfficeXmlNode DrawingProperties => _frame.Child(Xdr, "nvGraphicFramePr").Child(Xdr, "cNvPr");

        /// <summary>Non-visual drawing name.</summary>
        public string Name {
            get => Locking.ExecuteRead(_sheet.Document.EnsureLock(), () => {
                EnsureAttached();
                return DrawingProperties.Attribute("name") ?? string.Empty;
            });
            set {
                if (string.IsNullOrWhiteSpace(value)) throw new ArgumentNullException(nameof(value));
                ExcelSheet.ValidateModernChartText(value, nameof(value));
                Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                    EnsureAttached();
                    OfficeXmlNode properties = DrawingProperties;
                    if (properties.IsNone) throw new InvalidOperationException("Modern chart drawing properties are missing.");
                    string resolved = value.Trim();
                    OfficeXmlNode root = _drawingsPart.Root;
                    bool duplicate = root.Exists && root.Descendants(Xdr, "cNvPr")
                        .Any(candidate => candidate != properties
                            && string.Equals(candidate.Attribute("name"), resolved, StringComparison.OrdinalIgnoreCase));
                    if (duplicate) {
                        throw new InvalidOperationException($"Worksheet drawing name '{resolved}' is already in use.");
                    }
                    properties.SetAttribute("name", resolved);
                    SaveDrawing();
                });
            }
        }

        /// <summary>Detected native ChartEx layout.</summary>
        public ExcelModernChartType ChartType {
            get => Locking.ExecuteRead(_sheet.Document.EnsureLock(), () => {
                EnsureAttached();
                string? relationshipId = GetChartRelationshipId(_frame);
                ExcelPartRef part = relationshipId == null ? default : TryGetChartPart(relationshipId);
                OfficeXmlNode series = part.Exists ? FirstDescendant(part.Root, "series") : default;
                return ExcelSheet.ParseModernChartType(series.IsNone ? null : series.Attribute("layoutId"));
            });
        }

        /// <summary>Known authored data range, when this wrapper created the chart.</summary>
        public ExcelChartDataRange? DataRange { get; private set; }

        /// <summary>Current title text, when present.</summary>
        public string? Title => Locking.ExecuteRead(_sheet.Document.EnsureLock(), () => {
            OfficeXmlNode title = FirstDescendant(GetChartPart().Root, "title");
            OfficeXmlNode text = title.IsNone ? default : FirstDescendant(title, "v");
            return text.IsNone ? null : text.InnerText;
        });

        /// <summary>Sets the chart title while preserving unrelated imported ChartEx markup.</summary>
        public ExcelModernChart SetTitle(string? title) {
            ExcelSheet.ValidateModernChartText(title, nameof(title));
            return Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                ExcelPartRef part = GetChartPart();
                OfficeXmlNode root = part.Root;
                if (root.IsNone) throw new InvalidOperationException("Modern chart root is missing.");
                OfficeXmlNode chart = root.Child(ChartExUri, "chart");
                if (chart.IsNone) throw new InvalidOperationException("Modern chart content is missing.");
                RemoveChildren(chart, "title");
                if (!string.IsNullOrWhiteSpace(title)) {
                    OfficeXmlNode chartTitle = chart.CreateElement(ChartExUri, "title", "cx");
                    OfficeXmlNode plotArea = chart.Child(ChartExUri, "plotArea");
                    chartTitle = chart.InsertBefore(chartTitle, plotArea);
                    OfficeXmlNode tx = chartTitle.Append(chartTitle.CreateElement(ChartExUri, "tx", "cx"));
                    OfficeXmlNode txData = tx.Append(tx.CreateElement(ChartExUri, "txData", "cx"));
                    OfficeXmlNode v = txData.Append(txData.CreateElement(ChartExUri, "v", "cx"));
                    v.SetInnerText(title!.Trim());
                }
                part.SaveRoot();
                _sheet.MarkRequiresSavePreparation();
                return this;
            });
        }

        /// <summary>Changes the native series layout while preserving data, formatting, and unknown siblings.</summary>
        public ExcelModernChart SetChartType(ExcelModernChartType chartType) {
            string layout = ExcelSheet.GetModernChartLayout(chartType);
            return Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                ExcelPartRef part = GetChartPart();
                List<OfficeXmlNode> series = Descendants(part.Root, "series");
                if (series.Count == 0) throw new InvalidOperationException("Modern chart has no series to mutate.");
                foreach (OfficeXmlNode item in series) {
                    item.SetAttribute("layoutId", layout);
                }
                part.SaveRoot();
                _sheet.MarkRequiresSavePreparation();
                return this;
            });
        }

        /// <summary>Updates data for an OfficeIMO-authored chart while preserving unrelated ChartEx siblings and formatting.</summary>
        public ExcelModernChart UpdateData(ExcelChartData data) {
            ArgumentNullException.ThrowIfNull(data);
            return Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                ExcelChartDataRange currentRange = DataRange
                    ?? throw new InvalidOperationException("Imported ChartEx data cannot be replaced without an explicit authored data range.");
                if (data.Series.Count == 0 || data.Categories.Count == 0) {
                    throw new ArgumentException("Modern charts require at least one category and one series.", nameof(data));
                }
                ExcelSheet.ValidateModernChartData(data);
                ExcelPartRef part = GetChartPart();
                OfficeXmlNode root = part.Root;
                if (root.IsNone) throw new InvalidOperationException("Modern chart root is missing.");
                OfficeXmlNode chart = root.Child(ChartExUri, "chart");
                if (chart.IsNone) throw new InvalidOperationException("Modern chart content is missing.");
                OfficeXmlNode plotArea = chart.Child(ChartExUri, "plotArea");
                if (plotArea.IsNone) throw new InvalidOperationException("Modern chart plot area is missing.");
                OfficeXmlNode currentRegion = plotArea.Child(ChartExUri, "plotAreaRegion");
                string layout = ExcelSheet.GetModernChartLayout(ChartType);
                int pointCount = data.Categories.Count;
                ExcelChartDataRange updatedRange = currentRange.WithSize(pointCount, data.Series.Count);
                ExcelSheet dataSheet = _sheet.Document.GetSheetForLockedOperation(updatedRange.SheetName);
                bool sharedRange = _sheet.Document.IsOwnedChartDataRangeShared(currentRange, part);
                if (sharedRange
                    || pointCount > currentRange.CategoryCount
                    || data.Series.Count > currentRange.SeriesCount) {
                    int startRow = _sheet.Document.ReserveChartDataStartRow(dataSheet, pointCount + 1);
                    updatedRange = dataSheet.WriteChartData(
                        data,
                        startRow,
                        currentRange.StartColumn,
                        includeHeaderRow: currentRange.HasHeaderRow,
                        orientation: currentRange.Orientation);
                } else {
                    dataSheet.WriteChartData(
                        data,
                        updatedRange.StartRow,
                        updatedRange.StartColumn,
                        includeHeaderRow: updatedRange.HasHeaderRow,
                        orientation: updatedRange.Orientation);
                }

                byte[] markup = System.Text.Encoding.UTF8.GetBytes(ExcelSheet.BuildModernChartSpaceMarkup(data, updatedRange, layout, Title));
                OfficeXmlTree parsed = OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());
                OfficeXmlNode replacement = new OfficeXmlNode(parsed, parsed.Root);
                OfficeXmlNode replacementData = replacement.Child(ChartExUri, "chartData");
                OfficeXmlNode currentData = root.Child(ChartExUri, "chartData");
                if (currentData.IsNone) root.Append(replacementData.Clone());
                else currentData.ReplaceWith(replacementData.Clone());

                OfficeXmlNode replacementRegion = FirstDescendant(replacement, "plotAreaRegion");
                if (currentRegion.IsNone) {
                    plotArea.Append(replacementRegion.Clone());
                } else {
                    MergePlotAreaSeries(currentRegion, replacementRegion);
                }
                part.SaveRoot();
                _sheet.MarkRequiresSavePreparation();
                DataRange = updatedRange;
                _sheet.Document.ReleaseOwnedChartDataRange(
                    currentRange,
                    sharedRange ? currentRange : updatedRange);
                return this;
            });
        }

        private static readonly OfficeXmlChildOrder OneCellAnchorOrder = new(
            new[] { Xdr, Xdr, Xdr, Xdr, Xdr, Xdr, Xdr, Xdr, Xdr },
            new[] { "from", "ext", "sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart", "clientData" },
            new[] { 0, 1, 2, 2, 2, 2, 2, 2, 3 });

        /// <summary>Moves and resizes the one-cell chart anchor.</summary>
        public ExcelModernChart SetPlacement(int row, int column, OfficeLength width, OfficeLength height) {
            ExcelSheet.ValidateModernChartPlacement(row, column, width, height);
            return Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                EnsureAttached();
                OfficeXmlNode anchor = _frame.Ancestor(Xdr, "oneCellAnchor");
                if (anchor.IsNone) throw new InvalidOperationException("Only one-cell modern chart anchors can be repositioned.");
                // Inserted empty and filled in place, so its children carry no namespace declarations of their own.
                OfficeXmlNode from = anchor.SetOrdered(anchor.New(Xdr, "from"), OneCellAnchorOrder);
                from.Append(from.NewText(Xdr, "col", (column - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)));
                from.Append(from.NewText(Xdr, "colOff", "0"));
                from.Append(from.NewText(Xdr, "row", (row - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)));
                from.Append(from.NewText(Xdr, "rowOff", "0"));
                long widthEmu = width.ToEmus();
                long heightEmu = height.ToEmus();
                anchor.SetOrdered(anchor.New(Xdr, "ext").Attr("cx", widthEmu).Attr("cy", heightEmu), OneCellAnchorOrder);
                OfficeXmlNode transform = _frame.Child(Xdr, "xfrm");
                if (transform.Exists) {
                    transform.SetOrdered(transform.New(A, "ext").Attr("cx", widthEmu).Attr("cy", heightEmu), TransformOrder);
                }
                SaveDrawing();
                return this;
            });
        }

        private static readonly OfficeXmlChildOrder TransformOrder = new(new[] { A, A }, new[] { "off", "ext" }, new[] { 0, 1 });

        /// <summary>Removes this chart drawing and its owned ChartEx part.</summary>
        public void Remove() {
            Locking.ExecuteWrite(_sheet.Document.EnsureLock(), () => {
                string relationshipId = GetChartRelationshipId();
                ExcelPartRef part = GetChartPart(relationshipId);
                OfficeXmlNode anchor = _frame.Parent;
                if (anchor.Exists) anchor.Remove();
                OfficeXmlNode root = _drawingsPart.Root;
                bool partStillReferenced = root.Exists && root.Descendants(Xdr, "graphicFrame")
                    .Any(frame => string.Equals(GetChartRelationshipId(frame), relationshipId, StringComparison.Ordinal));
                if (!partStillReferenced) {
                    _drawingsPart.DeletePart(part);
                    if (DataRange != null) _sheet.Document.ReleaseOwnedChartDataRange(DataRange);
                }
                if (root.Exists && root.FirstChild.Exists) {
                    SaveDrawing();
                    return;
                }
                ExcelPartRef worksheetPart = ExcelPartRef.Worksheet(_sheet);
                OfficeXmlNode drawing = ExcelPartRef.WorksheetDrawingElement(_sheet);
                if (drawing.Exists) drawing.Remove();
                worksheetPart.DeletePart(_drawingsPart);
                ExcelPartRef.SaveWorksheetElement(_sheet);
                _sheet.MarkRequiresSavePreparation();
            });
        }

        private static void MergePlotAreaSeries(OfficeXmlNode currentRegion, OfficeXmlNode replacementRegion) {
            List<OfficeXmlNode> currentSeries = ChildrenNamed(currentRegion, "series");
            List<OfficeXmlNode> replacementSeries = ChildrenNamed(replacementRegion, "series");
            int retainedCount = Math.Min(currentSeries.Count, replacementSeries.Count);
            for (int index = 0; index < retainedCount; index++) {
                CopySeriesAttribute(currentSeries[index], replacementSeries[index], "layoutId");
                CopySeriesAttribute(currentSeries[index], replacementSeries[index], "ownerIdx");
                ReplaceSeriesChild(currentSeries[index], replacementSeries[index], "tx");
                ReplaceSeriesChild(currentSeries[index], replacementSeries[index], "dataId");
            }

            for (int index = replacementSeries.Count; index < currentSeries.Count; index++) {
                currentSeries[index].Remove();
            }

            OfficeXmlNode insertionPoint = currentSeries.LastOrDefault(series => series.Parent.Exists);
            for (int index = currentSeries.Count; index < replacementSeries.Count; index++) {
                OfficeXmlNode clone = replacementSeries[index].Clone();
                insertionPoint = insertionPoint.IsNone
                    ? currentRegion.InsertBefore(clone, currentRegion.FirstChild)
                    : insertionPoint.InsertAfterSelf(clone);
            }
        }

        private static void CopySeriesAttribute(OfficeXmlNode current, OfficeXmlNode replacement, string localName) {
            string value = replacement.Attribute(localName) ?? throw new InvalidOperationException("Sequence contains no matching element");
            current.SetAttribute(localName, value);
        }

        private static void ReplaceSeriesChild(OfficeXmlNode current, OfficeXmlNode replacement, string localName) {
            OfficeXmlNode replacementChild = default;
            foreach (OfficeXmlNode child in replacement.Children()) {
                if (string.Equals(child.LocalName, localName, StringComparison.Ordinal)) { replacementChild = child; break; }
            }
            if (replacementChild.IsNone) throw new InvalidOperationException("Sequence contains no matching element");
            OfficeXmlNode currentChild = default;
            foreach (OfficeXmlNode child in current.Children()) {
                if (string.Equals(child.LocalName, localName, StringComparison.Ordinal)) { currentChild = child; break; }
            }
            OfficeXmlNode clone = replacementChild.Clone();
            if (currentChild.IsNone) current.Append(clone);
            else currentChild.ReplaceWith(clone);
        }

        private ExcelPartRef GetChartPart() => GetChartPart(GetChartRelationshipId());

        private ExcelPartRef GetChartPart(string relationshipId) {
            EnsureAttached();
            ExcelPartRef part = TryGetChartPart(relationshipId);
            return part.Exists ? part : throw new InvalidOperationException("Modern chart part is missing.");
        }

        private ExcelPartRef TryGetChartPart(string relationshipId) {
            foreach ((string id, ExcelPartRef part) in _drawingsPart.Parts) {
                if (string.Equals(id, relationshipId, StringComparison.Ordinal)) {
                    return part.Kind == ExcelPartKind.ExtendedChart ? part : default;
                }
            }
            return default;
        }

        private void EnsureAttached() {
            OfficeXmlNode root = _drawingsPart.Root;
            if (_frame.Parent.IsNone
                || root.IsNone
                || !root.Descendants(Xdr, "graphicFrame").Any(frame => frame == _frame)) {
                throw new InvalidOperationException("Modern chart is no longer attached to the worksheet drawing.");
            }
        }

        private string GetChartRelationshipId() => GetChartRelationshipId(_frame)
            ?? throw new InvalidOperationException("Modern chart relationship identifier is missing.");

        private static string? GetChartRelationshipId(OfficeXmlNode frame) {
            OfficeXmlNode graphicData = frame.Child(A, "graphic").Child(A, "graphicData");
            OfficeXmlNode reference = graphicData.Child(ChartExUri, "chart");
            return reference.IsNone ? null : reference.Attribute(R, "id");
        }

        private static OfficeXmlNode FirstDescendant(OfficeXmlNode element, string localName) {
            if (element.IsNone) return default;
            for (OfficeXmlNode current = element.FirstChild; current.Exists; current = current.NextInSubtree(element)) {
                if (current.Is(ChartExUri, localName)) return current;
            }
            return default;
        }

        private static List<OfficeXmlNode> Descendants(OfficeXmlNode element, string localName) {
            var found = new List<OfficeXmlNode>();
            CollectDescendants(element, localName, found);
            return found;
        }

        private static void CollectDescendants(OfficeXmlNode element, string localName, List<OfficeXmlNode> found) {
            if (element.IsNone) return;
            for (OfficeXmlNode current = element.FirstChild; current.Exists; current = current.NextInSubtree(element)) {
                if (current.Is(ChartExUri, localName)) found.Add(current);
            }
        }

        private static List<OfficeXmlNode> ChildrenNamed(OfficeXmlNode parent, string localName) {
            var found = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in parent.Children()) {
                if (child.Is(ChartExUri, localName)) found.Add(child);
            }
            return found;
        }

        private static void RemoveChildren(OfficeXmlNode parent, string localName) {
            for (OfficeXmlNode child = parent.FirstChild; child.Exists;) {
                OfficeXmlNode next = child.NextSibling;
                if (child.Is(ChartExUri, localName)) child.Remove();
                child = next;
            }
        }

        private void SaveDrawing() {
            _drawingsPart.SaveRoot();
            _sheet.MarkRequiresSavePreparation();
        }
    }
}
