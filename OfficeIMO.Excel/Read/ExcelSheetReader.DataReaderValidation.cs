#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private bool ValidateDataReaderProjection(CancellationToken ct) {
            if (_canStreamWorksheetPart) {
                try {
                    using var stream = OpenDataReaderWorksheetStream(ct);
                    if (TryPrepareWorksheetStream(stream)) {
                        return ValidateDataReaderProjectionXml(stream, ct);
                    }
                } catch (XmlException) {
                } catch (IOException) {
                } catch (UnauthorizedAccessException) {
                } catch (ObjectDisposedException) {
                }
            }

            RequireSdkWorksheetPart();
            ct.ThrowIfCancellationRequested();
            Worksheet worksheet = _wsPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{_sheetName}' has no worksheet root.");
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                ct.ThrowIfCancellationRequested();
                var reference = new XmlCoordinateReference(cell.CellReference?.Value);
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
                    $"'{_sheetName}'!{reference.ToString()}. Read the workbook through ExcelDocument when resolved " +
                    "shared-formula text is required.");
            }
            return false;
        }

        private bool ValidateDataReaderProjectionXml(
            Stream stream,
            CancellationToken ct) {
            using var reader = OpenWorksheetXmlReader(stream);
            // Table-backed dimensions can intentionally include empty cells.
            // Otherwise reuse this complete validation scan to discover actual bounds.
            var bounds = _usedRangeA1 == null && (!_hasSdkWorksheetPart || !_wsPart.TableDefinitionParts.Any())
                ? new WorksheetRangeAccumulator() : null;
            var coordinates = bounds != null && Volatile.Read(ref _implicitXmlRowIndexes) == null
                ? new ImplicitXmlRowIndexBuilder() : null;
            bool haveSheetData = false;
            bool completedSheetData = false;
            int sheetDataDepth = -1;
            int rowDepth = -1;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                XmlNodeType nodeType = reader.NodeType;
                string localName = reader.LocalName;
                if (nodeType != XmlNodeType.Element || localName != "c") {
                    if (bounds != null) {
                        bool spreadsheetElement = reader.NamespaceURI == SpreadsheetNamespace
                            || reader.NamespaceURI == StrictSpreadsheetNamespace;
                        if (nodeType == XmlNodeType.Element && reader.Depth == 0
                            && (localName != "worksheet" || !spreadsheetElement)) {
                            bounds = null;
                        } else if (!_hasSdkWorksheetPart && localName == "tableParts"
                            && nodeType == XmlNodeType.Element) {
                            // Table extents may include intentional empty rows and
                            // columns; their discovery retains the SDK owner.
                            bounds = null;
                        } else if (localName == "sheetData" && nodeType == XmlNodeType.Element) {
                            if (haveSheetData || reader.Depth != 1 || !spreadsheetElement) {
                                bounds = null;
                            } else {
                                haveSheetData = true;
                                completedSheetData = reader.IsEmptyElement;
                                sheetDataDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            }
                        } else if (localName == "row" && nodeType == XmlNodeType.Element) {
                            if (sheetDataDepth < 0 || rowDepth >= 0
                                || reader.Depth != sheetDataDepth + 1 || !spreadsheetElement) {
                                bounds = null;
                            } else {
                                int declaredRowIndex = bounds.BeginRow(ReadXmlReferenceAttribute(reader).Text);
                                coordinates?.BeginRow(reader, declaredRowIndex);
                                if (reader.IsEmptyElement) {
                                    bounds.EndRow();
                                    coordinates?.EndRow();
                                } else rowDepth = reader.Depth;
                            }
                        } else if (nodeType == XmlNodeType.EndElement) {
                            if (reader.Depth == rowDepth && localName == "row") {
                                bounds.EndRow();
                                coordinates?.EndRow();
                                rowDepth = -1;
                            } else if (reader.Depth == sheetDataDepth && localName == "sheetData") {
                                completedSheetData = true;
                                sheetDataDepth = -1;
                            }
                        }
                    }
                    continue;
                }

                XmlCoordinateReference reference = ReadXmlReferenceAttribute(reader);
                if (bounds != null) {
                    if (rowDepth >= 0 && reader.Depth == rowDepth + 1
                        && (reader.NamespaceURI == SpreadsheetNamespace || reader.NamespaceURI == StrictSpreadsheetNamespace)) {
                        bounds.AddCell(reference.Text);
                        coordinates?.AddCell(reference.Text);
                    } else {
                        bounds = null;
                    }
                }
                XmlStyleAttribute styleIndex = ReadXmlStyleAttribute(reader);
                if (styleIndex.Present) {
                    ValidateCellStyleReference(styleIndex, reference);
                }
                bool sharedStringCell = string.Equals(
                    ReadXmlCellTypeAttribute(reader),
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
                    $"'{_sheetName}'!{reference.ToString()}. Read the workbook through ExcelDocument when resolved " +
                    "shared-formula text is required.");
            }
            ct.ThrowIfCancellationRequested();
            if (bounds != null && completedSheetData && rowDepth < 0) {
                if (bounds.TryGetReference(out string usedRangeReference)) {
                    _usedRangeA1 = usedRangeReference;
                }
                if (coordinates != null) {
                    Interlocked.CompareExchange(ref _implicitXmlRowIndexes, coordinates.Indexes, null);
                    return coordinates.RowsStrictlyIncreasing;
                }
            }
            return false;
        }

        private static string ReadSimpleElementText(
            XmlReader reader,
            CancellationToken ct,
            XmlCoordinateReference cellReference) {
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
                    $"Worksheet cell {cellReference.ToString()} element '{elementName}' must contain only text.");
            }

            return value;
        }

        private void ValidateCellStyleReference(XmlStyleAttribute style, XmlCoordinateReference reference) {
            if (style.Valid) {
                ValidateCellStyleReference(style.Index, reference);
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} contains an invalid cell style index.");
        }

        private void ValidateCellStyleReference(uint styleIndex, XmlCoordinateReference reference) {
            if (styleIndex < (uint)Styles.CellFormatCount) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} references a missing cell style.");
        }

        private void ValidateSharedStringReference(string? rawIndex, XmlCoordinateReference reference) {
            var items = _sharedStringItems ??= _sst.GetItems();
            if (TryParseSharedStringIndex(rawIndex, out int index)
                && (uint)index < (uint)items.Count) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} references a missing shared string.");
        }

    }
}
