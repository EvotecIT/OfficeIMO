#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        internal void ValidateDataReaderProjection(CancellationToken ct) {
            if (_canStreamWorksheetPart) {
                try {
                    using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (TryPrepareWorksheetStream(stream)) {
                        ValidateDataReaderProjectionXml(stream, ct);
                        return;
                    }
                } catch (XmlException) {
                } catch (IOException) {
                } catch (UnauthorizedAccessException) {
                } catch (ObjectDisposedException) {
                }
            }

            ct.ThrowIfCancellationRequested();
            Worksheet worksheet = _wsPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{_sheetName}' has no worksheet root.");
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                ct.ThrowIfCancellationRequested();
                string reference = cell.CellReference?.Value ?? "(unknown cell)";
                if (cell.StyleIndex?.Value is uint styleIndex) {
                    ValidateCellStyleReference(styleIndex, reference);
                }
                if (cell.DataType?.Value == CellValues.SharedString) {
                    ValidateSharedStringReference(cell.CellValue?.Text, reference);
                }

                CellFormula? formula = cell.CellFormula;
                if (formula?.FormulaType?.Value != CellFormulaValues.Shared
                    || !string.IsNullOrWhiteSpace(formula?.Text)) {
                    continue;
                }
                if (_opt.UseCachedFormulaResult && cell.CellValue is not null) {
                    continue;
                }

                throw new NotSupportedException(
                    $"Data-reader projection cannot safely expand the shared-formula follower " +
                    $"'{_sheetName}'!{reference}. Read the workbook through ExcelDocument when resolved " +
                    "shared-formula text is required.");
            }
        }

        private void ValidateDataReaderProjectionXml(
            Stream stream,
            CancellationToken ct) {
            using var reader = OpenWorksheetXmlReader(stream);
            // Table-backed dimensions can intentionally include empty cells.
            // Otherwise reuse this complete validation scan to discover actual bounds.
            var bounds = _usedRangeA1 == null && !_wsPart.TableDefinitionParts.Any()
                ? new WorksheetRangeAccumulator() : null;
            bool haveSheetData = false;
            bool completedSheetData = false;
            int sheetDataDepth = -1;
            int rowDepth = -1;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (bounds != null) {
                    bool spreadsheetElement = reader.NamespaceURI == SpreadsheetNamespace
                        || reader.NamespaceURI == StrictSpreadsheetNamespace;
                    if (reader.NodeType == XmlNodeType.Element && reader.Depth == 0
                        && (reader.LocalName != "worksheet" || !spreadsheetElement)) {
                        bounds = null;
                    } else if (reader.LocalName == "sheetData" && reader.NodeType == XmlNodeType.Element) {
                        if (haveSheetData || reader.Depth != 1 || !spreadsheetElement) {
                            bounds = null;
                        } else {
                            haveSheetData = true;
                            completedSheetData = reader.IsEmptyElement;
                            sheetDataDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                        }
                    } else if (reader.LocalName == "row" && reader.NodeType == XmlNodeType.Element) {
                        if (sheetDataDepth < 0 || rowDepth >= 0
                            || reader.Depth != sheetDataDepth + 1 || !spreadsheetElement) {
                            bounds = null;
                        } else {
                            bounds.BeginRow(reader.GetAttribute("r"));
                            if (reader.IsEmptyElement) bounds.EndRow();
                            else rowDepth = reader.Depth;
                        }
                    } else if (reader.NodeType == XmlNodeType.EndElement) {
                        if (reader.Depth == rowDepth && reader.LocalName == "row") {
                            bounds.EndRow();
                            rowDepth = -1;
                        } else if (reader.Depth == sheetDataDepth && reader.LocalName == "sheetData") {
                            completedSheetData = true;
                            sheetDataDepth = -1;
                        }
                    }
                }
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                    continue;
                }

                string? rawReference = reader.GetAttribute("r");
                if (bounds != null) {
                    if (rowDepth >= 0 && reader.Depth == rowDepth + 1
                        && (reader.NamespaceURI == SpreadsheetNamespace || reader.NamespaceURI == StrictSpreadsheetNamespace)) {
                        bounds.AddCell(rawReference);
                    } else {
                        bounds = null;
                    }
                }
                string reference = rawReference ?? "(unknown cell)";
                string? styleIndex = reader.GetAttribute("s");
                if (styleIndex != null) {
                    ValidateCellStyleReference(styleIndex, reference);
                }
                bool sharedStringCell = string.Equals(
                    reader.GetAttribute("t"),
                    "s",
                    StringComparison.Ordinal);
                bool sharedFollower = false;
                bool hasCachedValue = false;
                string? sharedStringReference = null;
                if (reader.IsEmptyElement) {
                    if (sharedStringCell) {
                        ValidateSharedStringReference(null, reference);
                    }
                    continue;
                }

                int cellDepth = reader.Depth;
                while (reader.Read()) {
                    ct.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.EndElement
                        && reader.Depth == cellDepth
                        && reader.LocalName == "c") {
                        break;
                    }
                    if (reader.NodeType != XmlNodeType.Element
                        || reader.Depth != cellDepth + 1
                        || (!string.Equals(
                                reader.NamespaceURI,
                                SpreadsheetNamespace,
                                StringComparison.Ordinal)
                            && !string.Equals(
                                reader.NamespaceURI,
                                StrictSpreadsheetNamespace,
                                StringComparison.Ordinal))) {
                        continue;
                    }

                    if (reader.LocalName == "v") {
                        hasCachedValue = true;
                        if (sharedStringCell) {
                            sharedStringReference = reader.IsEmptyElement
                                ? string.Empty
                                : ReadSimpleElementText(reader, ct, reference);
                        }
                        continue;
                    }

                    if (reader.LocalName != "f"
                        || !string.Equals(
                            reader.GetAttribute("t"),
                            "shared",
                            StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    bool isFollower = reader.IsEmptyElement;
                    if (!isFollower) {
                        isFollower = string.IsNullOrWhiteSpace(
                            ReadSimpleElementText(reader, ct, reference));
                    }
                    sharedFollower |= isFollower;
                }

                if (sharedStringCell) {
                    ValidateSharedStringReference(sharedStringReference, reference);
                }

                if (!sharedFollower
                    || (_opt.UseCachedFormulaResult && hasCachedValue)) {
                    continue;
                }

                throw new NotSupportedException(
                    $"Data-reader projection cannot safely expand the shared-formula follower " +
                    $"'{_sheetName}'!{reference}. Read the workbook through ExcelDocument when resolved " +
                    "shared-formula text is required.");
            }
            ct.ThrowIfCancellationRequested();
            if (bounds != null && completedSheetData && rowDepth < 0
                && bounds.TryGetReference(out string usedRangeReference)) {
                _usedRangeA1 = usedRangeReference;
            }
        }

        private static string ReadSimpleElementText(
            XmlReader reader,
            CancellationToken ct,
            string cellReference) {
            int elementDepth = reader.Depth;
            string elementName = reader.LocalName;
            string elementNamespace = reader.NamespaceURI;
            string value = reader.ReadString();
            ct.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.EndElement
                || reader.Depth != elementDepth
                || !string.Equals(reader.LocalName, elementName, StringComparison.Ordinal)
                || !string.Equals(reader.NamespaceURI, elementNamespace, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"Worksheet cell {cellReference} element '{elementName}' must contain only text.");
            }

            return value;
        }

        private void ValidateCellStyleReference(string rawIndex, string reference) {
            if (TryParseUInt(rawIndex, out uint styleIndex)) {
                ValidateCellStyleReference(styleIndex, reference);
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference} contains an invalid cell style index.");
        }

        private void ValidateCellStyleReference(uint styleIndex, string reference) {
            if (styleIndex < (uint)Styles.CellFormatCount) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference} references a missing cell style.");
        }

        private void ValidateSharedStringReference(string? rawIndex, string reference) {
            var items = _sharedStringItems ??= _sst.GetItems();
            if (TryParseSharedStringIndex(rawIndex, out int index)
                && (uint)index < (uint)items.Count) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference} references a missing shared string.");
        }

    }
}
