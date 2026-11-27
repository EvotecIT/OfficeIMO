using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>
        /// The chart part the frame refers to. Resolving it also refreshes a data range invalidated by a save or a sheet
        /// rename.
        /// </summary>
        private ExcelPartRef ChartPartRef() {
            string? relationshipId = FrameNode.Child(A, "graphic").Child(A, "graphicData").Child(C, "chart").RelationshipAttribute("id");
            if (relationshipId == null) {
                throw new InvalidOperationException("Chart reference not found for the shape.");
            }

            ExcelPartRef part = DrawingsPartRef.GetChartPartById(relationshipId);
            if (_dataRangeNeedsRefresh || !string.Equals(_lastKnownSheetName, _sheetName, StringComparison.Ordinal)) {
                if (_dataRange != null) _dataRange = ExcelChartUtils.TryExtractDataRange(part, _dataRange.SeriesCount);
                _lastKnownSheetName = _sheetName;
                _dataRangeNeedsRefresh = false;
            }
            return part;
        }

        /// <summary>The chart part's <c>c:chartSpace</c> root, or none.</summary>
        private OfficeXmlNode ChartSpaceNode() => ChartPartRef().Root;

        /// <summary>The <c>c:chart</c> element; throws when the chart part has none.</summary>
        private OfficeXmlNode ChartNode() {
            OfficeXmlNode chart = ChartPartRef().Root.Child(C, "chart");
            if (chart.IsNone) {
                throw new InvalidOperationException("Chart element not found in chart part.");
            }
            return chart;
        }
    }
}
