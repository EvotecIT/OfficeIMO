using System.Globalization;
using System.Xml;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// The parts streaming AutoFit reads from one serialized cell that has no parent row, rich text, formula or shared
        /// string: its style (0 when absent), its type, its value text and, for an inline string, its single text element.
        /// </summary>
        internal readonly struct AutoFitCellParts {
            internal AutoFitCellParts(uint styleIndex, XCellValues? dataType, string value, string? inlineText) {
                StyleIndex = styleIndex;
                DataType = dataType;
                Value = value;
                InlineText = inlineText;
            }

            internal uint StyleIndex { get; }
            internal XCellValues? DataType { get; }
            internal string Value { get; }
            internal string? InlineText { get; }

            /// <summary>How <see cref="Read"/> handled a cell.</summary>
            internal enum ReadResult {
                /// <summary>The parts were read and the reader is past the cell.</summary>
                Read,
                /// <summary>A shared string or unknown type; the reader has not moved, so the caller reads the cell's markup.</summary>
                Markup,
                /// <summary>Unexpected content was skipped; the caller measures the whole sheet from markup instead.</summary>
                Unsupported
            }

            /// <summary>
            /// Reads a streaming cell positioned at its start element: the <c>s</c> and <c>t</c> attributes and either nothing,
            /// a <c>v</c> element (after any <c>f</c> element, which measuring ignores) or an inline string with one <c>t</c> element.
            /// </summary>
            internal static ReadResult Read(XmlReader reader, out AutoFitCellParts parts) {
                parts = default;
                uint styleIndex = 0U;
                string? style = reader.GetAttribute("s");
                if (style != null && !uint.TryParse(style, NumberStyles.None, CultureInfo.InvariantCulture, out styleIndex)) {
                    return ReadResult.Markup;
                }

                XCellValues? dataType;
                switch (reader.GetAttribute("t")) {
                    case null: dataType = null; break;
                    case "n": dataType = XCellValues.Number; break;
                    case "b": dataType = XCellValues.Boolean; break;
                    case "e": dataType = XCellValues.Error; break;
                    case "str": dataType = XCellValues.String; break;
                    case "inlineStr": dataType = XCellValues.InlineString; break;
                    case "d": dataType = XCellValues.Date; break;
                    default: return ReadResult.Markup;
                }

                if (reader.IsEmptyElement) {
                    reader.Read();
                    parts = new AutoFitCellParts(styleIndex, dataType, string.Empty, null);
                    return ReadResult.Read;
                }

                int depth = reader.Depth;
                string value = string.Empty;
                string? inlineText = null;
                bool supported = true;
                reader.Read();
                while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) && !reader.EOF) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        reader.Read();
                    } else if (reader.LocalName == "v" && reader.Depth == depth + 1) {
                        value = reader.ReadElementContentAsString();
                    } else if (reader.LocalName == "f" && reader.Depth == depth + 1) {
                        reader.Skip();
                    } else if (reader.LocalName == "is" && reader.Depth == depth + 1 && inlineText == null) {
                        supported &= TryReadInlineText(reader, out inlineText);
                    } else {
                        supported = false;
                        reader.Skip();
                    }
                }

                reader.Read();
                if (!supported) return ReadResult.Unsupported;
                parts = new AutoFitCellParts(styleIndex, dataType, value, inlineText);
                return ReadResult.Read;
            }

            /// <summary>Reads an inline string holding one text element; false (after skipping it) for runs or other content.</summary>
            private static bool TryReadInlineText(XmlReader reader, out string? text) {
                text = string.Empty;
                if (reader.IsEmptyElement) {
                    reader.Read();
                    return true;
                }

                int depth = reader.Depth;
                reader.Read();
                while (reader.NodeType != XmlNodeType.Element && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
                    if (!reader.Read()) return false;
                }

                if (reader.NodeType == XmlNodeType.EndElement) {
                    reader.Read();
                    return true;
                }

                if (reader.LocalName != "t") {
                    SkipTo(reader, depth);
                    return false;
                }

                text = reader.ReadElementContentAsString();
                while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
                    if (reader.NodeType == XmlNodeType.Element) {
                        SkipTo(reader, depth);
                        return false;
                    }

                    if (!reader.Read()) return false;
                }

                reader.Read();
                return true;
            }

            private static void SkipTo(XmlReader reader, int depth) {
                while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) && !reader.EOF) {
                    if (reader.NodeType == XmlNodeType.Element) reader.Skip();
                    else reader.Read();
                }

                reader.Read();
            }
        }

        internal sealed partial class AutoFitWidthMeasurer {
            /// <summary>Measures one streaming cell from its parts exactly as <see cref="MeasureCell(DocumentFormat.OpenXml.Spreadsheet.Cell, ExcelDateSystem, bool, bool)"/> measures it.</summary>
            internal double MeasureCell(in AutoFitCellParts cell, ExcelDateSystem dateSystem, bool tableHeader, bool filterButton) {
                uint style = cell.StyleIndex;
                string text = TryGetDateAutoFitSample(cell.DataType, style, cell.Value, textContext, out string sample)
                    ? sample : GetAutoFitText(cell, dateSystem);
                double width = Measure(text, style);
                if (tableHeader) {
                    width = Math.Max(width, MeasureHeader(GetAutoFitText(cell, dateSystem), style, bold: true, filterButton));
                }

                return width;
            }

            private string GetAutoFitText(in AutoFitCellParts cell, ExcelDateSystem dateSystem) =>
                cell.DataType == XCellValues.InlineString
                    ? FormatAutoFitTextValue(Utilities.ExcelSanitizer.DecodeXString(cell.InlineText ?? string.Empty), cell.StyleIndex, textContext)
                    : ExcelSheet.GetAutoFitText(cell.DataType, cell.StyleIndex, cell.Value, textContext, dateSystem);
        }
    }
}
