using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

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
        private readonly Xdr.GraphicFrame _frame;
        private readonly DrawingsPart _drawingsPart;
        private readonly ExcelSheet _sheet;

        internal ExcelModernChart(
            Xdr.GraphicFrame frame,
            DrawingsPart drawingsPart,
            ExcelSheet sheet,
            ExcelChartDataRange? dataRange = null) {
            _frame = frame;
            _drawingsPart = drawingsPart;
            _sheet = sheet;
            DataRange = dataRange;
        }

        /// <summary>Non-visual drawing name.</summary>
        public string Name {
            get => _frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            set {
                if (string.IsNullOrWhiteSpace(value)) throw new ArgumentNullException(nameof(value));
                EnsureAttached();
                Xdr.NonVisualDrawingProperties properties = _frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties
                    ?? throw new InvalidOperationException("Modern chart drawing properties are missing.");
                properties.Name = value.Trim();
                SaveDrawing();
            }
        }

        /// <summary>Detected native ChartEx layout.</summary>
        public ExcelModernChartType ChartType => ExcelSheet.ParseModernChartType(GetSeries().FirstOrDefault()?.GetAttribute("layoutId", string.Empty).Value);

        /// <summary>Known authored data range, when this wrapper created the chart.</summary>
        public ExcelChartDataRange? DataRange { get; private set; }

        /// <summary>Current title text, when present.</summary>
        public string? Title => GetChartPart().ChartSpace?.Descendants<Cx.ChartTitle>()
            .FirstOrDefault()?.Descendants<Cx.VXsdstring>().FirstOrDefault()?.Text;

        /// <summary>Sets the chart title while preserving unrelated imported ChartEx markup.</summary>
        public ExcelModernChart SetTitle(string? title) {
            ExtendedChartPart part = GetChartPart();
            Cx.ChartSpace root = part.ChartSpace ?? throw new InvalidOperationException("Modern chart root is missing.");
            Cx.Chart chart = root.GetFirstChild<Cx.Chart>()
                ?? throw new InvalidOperationException("Modern chart content is missing.");
            chart.RemoveAllChildren<Cx.ChartTitle>();
            if (!string.IsNullOrWhiteSpace(title)) {
                var chartTitle = new Cx.ChartTitle(
                    new Cx.Text(
                        new Cx.TextData(
                            new Cx.VXsdstring(title!.Trim()))));
                Cx.PlotArea? plotArea = chart.GetFirstChild<Cx.PlotArea>();
                if (plotArea == null) chart.Append(chartTitle);
                else chart.InsertBefore(chartTitle, plotArea);
            }
            part.ChartSpace.Save();
            _sheet.MarkRequiresSavePreparation();
            return this;
        }

        /// <summary>Changes the native series layout while preserving data, formatting, and unknown siblings.</summary>
        public ExcelModernChart SetChartType(ExcelModernChartType chartType) {
            string layout = ExcelSheet.GetModernChartLayout(chartType);
            Cx.Series[] series = GetSeries();
            if (series.Length == 0) throw new InvalidOperationException("Modern chart has no series to mutate.");
            foreach (Cx.Series item in series) {
                item.SetAttribute(new OpenXmlAttribute("layoutId", string.Empty, layout));
            }
            GetChartPart().ChartSpace!.Save();
            _sheet.MarkRequiresSavePreparation();
            return this;
        }

        /// <summary>Updates data for an OfficeIMO-authored chart while preserving unrelated ChartEx siblings and formatting.</summary>
        public ExcelModernChart UpdateData(ExcelChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            ExcelChartDataRange currentRange = DataRange
                ?? throw new InvalidOperationException("Imported ChartEx data cannot be replaced without an explicit authored data range.");
            if (data.Series.Count == 0 || data.Categories.Count == 0) {
                throw new ArgumentException("Modern charts require at least one category and one series.", nameof(data));
            }
            ExcelSheet.ValidateModernChartData(data);
            ExtendedChartPart part = GetChartPart();
            string layout = ExcelSheet.GetModernChartLayout(ChartType);
            int pointCount = data.Categories.Count;
            ExcelChartDataRange updatedRange = currentRange.WithSize(pointCount, data.Series.Count);
            ExcelSheet dataSheet = _sheet.Document[updatedRange.SheetName];
            if (pointCount > currentRange.CategoryCount || data.Series.Count > currentRange.SeriesCount) {
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

            Cx.ChartSpace root = part.ChartSpace ?? throw new InvalidOperationException("Modern chart root is missing.");
            Cx.ChartSpace replacement = ExcelSheet.BuildModernChartSpace(
                data,
                updatedRange,
                layout,
                Title);
            Cx.ChartData replacementData = replacement.GetFirstChild<Cx.ChartData>()!;
            Cx.ChartData? currentData = root.GetFirstChild<Cx.ChartData>();
            if (currentData == null) root.Append((Cx.ChartData)replacementData.CloneNode(true));
            else root.ReplaceChild((Cx.ChartData)replacementData.CloneNode(true), currentData);

            Cx.PlotAreaRegion replacementRegion = replacement.Descendants<Cx.PlotAreaRegion>().First();
            Cx.PlotAreaRegion? currentRegion = root.Descendants<Cx.PlotAreaRegion>().FirstOrDefault();
            if (currentRegion == null) {
                Cx.PlotArea plotArea = root.Descendants<Cx.PlotArea>().FirstOrDefault()
                    ?? throw new InvalidOperationException("Modern chart plot area is missing.");
                plotArea.Append((Cx.PlotAreaRegion)replacementRegion.CloneNode(true));
            } else {
                currentRegion.Parent!.ReplaceChild((Cx.PlotAreaRegion)replacementRegion.CloneNode(true), currentRegion);
            }
            part.ChartSpace.Save();
            _sheet.MarkRequiresSavePreparation();
            DataRange = updatedRange;
            _sheet.Document.ReleaseOwnedChartDataRange(currentRange, updatedRange);
            return this;
        }

        /// <summary>Moves and resizes the one-cell chart anchor.</summary>
        public ExcelModernChart SetPlacement(int row, int column, int widthPixels, int heightPixels) {
            ExcelSheet.ValidateModernChartPlacement(row, column, widthPixels, heightPixels);
            EnsureAttached();
            Xdr.OneCellAnchor anchor = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()
                ?? throw new InvalidOperationException("Only one-cell modern chart anchors can be repositioned.");
            anchor.FromMarker = new Xdr.FromMarker(
                new Xdr.ColumnId((column - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((row - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new Xdr.RowOffset("0"));
            long width = (long)Math.Round(widthPixels * 9525D);
            long height = (long)Math.Round(heightPixels * 9525D);
            anchor.Extent = new Xdr.Extent { Cx = width, Cy = height };
            Xdr.Transform? transform = _frame.Transform;
            if (transform != null) transform.Extents = new DocumentFormat.OpenXml.Drawing.Extents { Cx = width, Cy = height };
            SaveDrawing();
            return this;
        }

        /// <summary>Removes this chart drawing and its owned ChartEx part.</summary>
        public void Remove() {
            string relationshipId = GetChartRelationshipId();
            ExtendedChartPart part = GetChartPart(relationshipId);
            OpenXmlElement? anchor = _frame.Parent;
            anchor?.Remove();
            bool partStillReferenced = _drawingsPart.WorksheetDrawing?.Descendants<Xdr.GraphicFrame>()
                .Any(frame => string.Equals(GetChartRelationshipId(frame), relationshipId, StringComparison.Ordinal)) == true;
            if (!partStillReferenced) {
                _drawingsPart.DeletePart(part);
                if (DataRange != null) _sheet.Document.ReleaseOwnedChartDataRange(DataRange);
            }
            if (_drawingsPart.WorksheetDrawing?.ChildElements.Any() == true) {
                SaveDrawing();
                return;
            }
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = _sheet.WorksheetPart.Worksheet
                ?? throw new InvalidOperationException("Worksheet root is missing.");
            DocumentFormat.OpenXml.Spreadsheet.Drawing? drawing = worksheet
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            drawing?.Remove();
            _sheet.WorksheetPart.DeletePart(_drawingsPart);
            worksheet.Save();
            _sheet.MarkRequiresSavePreparation();
        }

        private Cx.Series[] GetSeries() => GetChartPart().ChartSpace?.Descendants<Cx.Series>().ToArray()
            ?? Array.Empty<Cx.Series>();

        private ExtendedChartPart GetChartPart() => GetChartPart(GetChartRelationshipId());

        private ExtendedChartPart GetChartPart(string relationshipId) {
            EnsureAttached();
            return _drawingsPart.GetPartById(relationshipId) as ExtendedChartPart
                ?? throw new InvalidOperationException("Modern chart part is missing.");
        }

        private void EnsureAttached() {
            if (_frame.Parent == null
                || _drawingsPart.WorksheetDrawing?.Descendants<Xdr.GraphicFrame>()
                    .Any(frame => ReferenceEquals(frame, _frame)) != true) {
                throw new InvalidOperationException("Modern chart is no longer attached to the worksheet drawing.");
            }
        }

        private string GetChartRelationshipId() => GetChartRelationshipId(_frame)
            ?? throw new InvalidOperationException("Modern chart relationship identifier is missing.");

        private static string? GetChartRelationshipId(Xdr.GraphicFrame frame) {
            OpenXmlElement? reference = frame.Graphic?.GraphicData?.ChildElements.FirstOrDefault(element =>
                string.Equals(element.LocalName, "chart", StringComparison.Ordinal)
                && string.Equals(element.NamespaceUri, "http://schemas.microsoft.com/office/drawing/2014/chartex", StringComparison.Ordinal));
            return reference?.GetAttribute(
                "id",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
        }

        private void SaveDrawing() {
            _drawingsPart.WorksheetDrawing?.Save();
            _sheet.MarkRequiresSavePreparation();
        }
    }
}
