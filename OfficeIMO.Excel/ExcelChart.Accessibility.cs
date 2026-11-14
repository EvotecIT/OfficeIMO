using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>Gets the chart's alternative text (the drawing description), or an empty string.</summary>
        public string AltText => DrawingProperties?.Description?.Value ?? string.Empty;

        /// <summary>Gets the chart's accessibility title, or an empty string.</summary>
        public string AltTextTitle => DrawingProperties?.Title?.Value ?? string.Empty;

        /// <summary>Gets whether the chart is marked decorative, so screen readers and accessibility checks skip it.</summary>
        public bool IsDecorative => ExcelDrawingAccessibility.IsDecorative(DrawingProperties);

        /// <summary>Sets the chart's alternative text and optional title, and clears any decorative flag.</summary>
        public ExcelChart SetAltText(string description, string? title = null) {
            Xdr.NonVisualDrawingProperties? properties = DrawingProperties;
            if (properties != null) {
                properties.Description = description ?? string.Empty;
                ExcelDrawingAccessibility.ClearDecorative(properties);
                if (title != null) properties.Title = title;
                SaveDrawing();
            }
            return this;
        }

        /// <summary>Marks the chart as decorative: clears its alternative text and sets the Office decorative flag.</summary>
        public ExcelChart Decorative() {
            Xdr.NonVisualDrawingProperties? properties = DrawingProperties;
            if (properties != null) {
                ExcelDrawingAccessibility.MarkDecorative(properties);
                SaveDrawing();
            }
            return this;
        }

        private Xdr.NonVisualDrawingProperties? DrawingProperties => _frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;

        private void SaveDrawing() {
            _drawingsPart.WorksheetDrawing?.Save();
            _document.MarkPackageDirty();
        }
    }
}
