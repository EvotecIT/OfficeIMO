using System.Globalization;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Style inheritance for displayed-text reads. A cell without its own <c>s</c> attribute displays with its row's
    /// custom style, otherwise with its column's style, as <see cref="ExcelCell.GetFormattedText"/> resolves it.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        private readonly object _displayColumnGate = new();
        private (uint Min, uint Max, uint Style)[]? _displayColumnStyles;

        // Row style of the row the sequential XML readers are positioned in; null when the row has no custom format.
        private uint? _displayRowStyle;

        /// <summary>True when any column of the sheet carries a style whose number format changes displayed text.</summary>
        private bool HasFormattedColumnStyles => DisplayColumnStyles.Length > 0;

        private (uint Min, uint Max, uint Style)[] DisplayColumnStyles {
            get {
                var styles = Volatile.Read(ref _displayColumnStyles);
                if (styles != null) return styles;
                lock (_displayColumnGate) {
                    return _displayColumnStyles ??= ReadFormattedColumnStyles();
                }
            }
        }

        /// <summary>Returns the inherited style for a cell without its own style, when that style changes its text.</summary>
        private uint? ResolveInheritedDisplayStyle(uint? rowStyle, int column) {
            ExcelCellDisplayText? display = _opt.DisplayText;
            if (display == null) return null;
            if (rowStyle.HasValue) return display.IsFormatted(rowStyle.Value) ? rowStyle : null;
            if (column <= 0) return null;
            foreach (var definition in DisplayColumnStyles) {
                if (definition.Min <= (uint)column && (uint)column <= definition.Max) return definition.Style;
            }
            return null;
        }

        /// <summary>Returns the one-based column of an A1 cell reference such as <c>C12</c>, or zero.</summary>
        private static int GetReferenceColumn(string? reference) {
            if (string.IsNullOrEmpty(reference)) return 0;
            int column = 0;
            foreach (char ch in reference!) {
                int letter = (ch | 0x20) - 'a';
                if ((uint)letter > 25U) break;
                column = column * 26 + letter + 1;
                if (column > ExcelA1.MaxColumns) return 0;
            }
            return column;
        }

        /// <summary>Captures the custom style of the row element <paramref name="rowReader"/> is positioned on.</summary>
        private void CaptureDisplayRowStyle(XmlReader rowReader) {
            if (_opt.DisplayText == null) return;
            _displayRowStyle = null;
            string? customFormat = rowReader.GetAttribute("customFormat");
            if (customFormat != "1" && !string.Equals(customFormat, "true", StringComparison.OrdinalIgnoreCase)) return;
            if (uint.TryParse(rowReader.GetAttribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out uint style)) {
                _displayRowStyle = style;
            }
        }

        /// <summary>Reads column definitions ahead of the cell data; only styles that change displayed text are kept.</summary>
        private (uint Min, uint Max, uint Style)[] ReadFormattedColumnStyles() {
            ExcelCellDisplayText? display = _opt.DisplayText;
            if (display == null) return Array.Empty<(uint, uint, uint)>();
            var styles = new List<(uint Min, uint Max, uint Style)>();
            if (_hasLiveWorksheetPart && _worksheetPart.IsRootLoaded) {
                // A loaded or imported worksheet lives in its DOM; its part stream may be stale.
                foreach (OfficeXmlNode columns in _worksheetPart.Root.Children(SpreadsheetXml, "cols")) {
                    foreach (OfficeXmlNode column in columns.Children(SpreadsheetXml, "col")) {
                        if (column.UInt32Attribute("min") is uint min && column.UInt32Attribute("max") is uint max
                            && column.UInt32Attribute("style") is uint style && display.IsFormatted(style)) {
                            styles.Add((min, max, style));
                        }
                    }
                }
                return styles.ToArray();
            }
            try {
                using Stream stream = OpenReaderWorksheetStream(_opt.CancellationToken);
                using XmlReader reader = OpenWorksheetXmlReader(stream);
                var worksheet = new WorksheetXmlRowSelector();
                int columnsDepth = -1;
                while (reader.Read()) {
                    _opt.CancellationToken.ThrowIfCancellationRequested();
                    if (worksheet.IsWorksheetChildElement(reader, "sheetData")) break;
                    if (reader.Depth <= columnsDepth) columnsDepth = -1;
                    if (worksheet.IsWorksheetChildElement(reader, "cols") && !reader.IsEmptyElement) {
                        columnsDepth = reader.Depth;
                        continue;
                    }
                    if (columnsDepth < 0 || !SpreadsheetXmlContent.IsDirectChildElement(reader, columnsDepth, "col")) continue;
                    if (uint.TryParse(reader.GetAttribute("min"), NumberStyles.None, CultureInfo.InvariantCulture, out uint min)
                        && uint.TryParse(reader.GetAttribute("max"), NumberStyles.None, CultureInfo.InvariantCulture, out uint max)
                        && uint.TryParse(reader.GetAttribute("style"), NumberStyles.None, CultureInfo.InvariantCulture, out uint style)
                        && display.IsFormatted(style)) {
                        styles.Add((min, max, style));
                    }
                }
            } catch (XmlException) {
                // The cell readers report malformed worksheets; inheritance simply stays unresolved here.
            }
            return styles.ToArray();
        }
    }
}
