using System.Collections.Generic;
using SixColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel.Fluent.Report
{
    public sealed class IconSetOptions
    {
        public DocumentFormat.OpenXml.Spreadsheet.IconSetValues IconSet { get; set; } = DocumentFormat.OpenXml.Spreadsheet.IconSetValues.ThreeTrafficLights1;
        public bool ShowValue { get; set; } = true;
        public bool ReverseOrder { get; set; } = false;
    }
    /// <summary>
    /// Declarative visual options for tables produced by ReportSheetBuilder.TableFrom.
    /// Keeps the API generic so callers can style by header names without project-specific code.
    /// </summary>
    public sealed class TableVisualOptions
    {
        /// <summary>Freeze through the header row for easier scrolling.</summary>
        public bool FreezeHeaderRow { get; set; } = true;

        /// <summary>
        /// Apply a numeric format with fixed decimals by header name.
        /// </summary>
        public Dictionary<string, int> NumericColumnDecimals { get; } = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Apply a custom number format pattern by header name.
        /// </summary>
        public Dictionary<string, string> NumericColumnFormats { get; } = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Add a data bar to a column with a specific color by header name.
        /// </summary>
        public Dictionary<string, SixColor> DataBars { get; } = new Dictionary<string, SixColor>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Add an icon set to columns identified by header name (simple).
        /// </summary>
        public HashSet<string> IconSetColumns { get; } = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Detailed icon set options by header name.
        /// </summary>
        public Dictionary<string, IconSetOptions> IconSets { get; } = new Dictionary<string, IconSetOptions>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Apply background colors depending on cell text for a given header name.
        /// </summary>
        public Dictionary<string, IDictionary<string, string>> TextBackgrounds { get; } = new Dictionary<string, IDictionary<string, string>>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// When true, applies generic numeric formatting and a subtle data bar to columns
        /// that originate from mapped collections (heuristic: path contains a '.').
        /// </summary>
        public bool AutoFormatDynamicCollections { get; set; } = true;

        /// <summary>
        /// Optional override for the data bar color used by AutoFormatDynamicCollections.
        /// </summary>
        public SixColor AutoFormatDataBarColor { get; set; } = SixColor.LightSkyBlue;

        /// <summary>
        /// Optional override for decimals used by AutoFormatDynamicCollections.
        /// </summary>
        public int AutoFormatDecimals { get; set; } = 2;
    }
}
