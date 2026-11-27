using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>Gets the chart's alternative text (the drawing description), or an empty string.</summary>
        public string AltText => DrawingProperties.Attribute("descr") ?? string.Empty;

        /// <summary>Gets the chart's accessibility title, or an empty string.</summary>
        public string AltTextTitle => DrawingProperties.Attribute("title") ?? string.Empty;

        /// <summary>Gets whether the chart is marked decorative, so screen readers and accessibility checks skip it.</summary>
        public bool IsDecorative => ExcelDrawingAccessibility.IsDecorative(DrawingProperties);

        /// <summary>Sets the chart's alternative text and optional title, and clears any decorative flag.</summary>
        public ExcelChart SetAltText(string description, string? title = null) {
            OfficeXmlNode properties = DrawingProperties;
            if (properties.Exists) {
                properties.SetAttribute("descr", description ?? string.Empty);
                ExcelDrawingAccessibility.ClearDecorative(properties);
                if (title != null) properties.SetAttribute("title", title);
                SaveDrawing();
            }
            return this;
        }

        /// <summary>Marks the chart as decorative: clears its alternative text and sets the Office decorative flag.</summary>
        public ExcelChart Decorative() {
            OfficeXmlNode properties = DrawingProperties;
            if (properties.Exists) {
                ExcelDrawingAccessibility.MarkDecorative(properties);
                SaveDrawing();
            }
            return this;
        }

        /// <summary>The frame's <c>xdr:cNvPr</c>, or none.</summary>
        private OfficeXmlNode DrawingProperties => FrameNode.Child(Xdr, "nvGraphicFramePr").Child(Xdr, "cNvPr");

        private void SaveDrawing() {
            DrawingsPartRef.SaveRoot();
            _document.MarkPackageDirty();
        }
    }
}
