using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private bool _reportedEmptyCellStyleRepair;

        private bool ValidateDataReaderProjection(CancellationToken ct) {
            if (_canStreamWorksheetPart) {
                try {
                    using var stream = OpenReaderWorksheetStream(ct);
                    if (TryPrepareWorksheetStream(stream)) {
                        return ValidateDataReaderProjectionXml(stream, ct);
                    }
                } catch (XmlException) {
                } catch (IOException exception) when (!ExcelReadLimitFailure.Is(exception)) {
                } catch (UnauthorizedAccessException) {
                } catch (ObjectDisposedException) {
                }
            }

            RequireGeneralProjection();
            ct.ThrowIfCancellationRequested();
            OfficeXmlNode worksheet = WorksheetRoot;
            if (worksheet.IsNone) throw new InvalidDataException($"Worksheet '{_sheetName}' has no worksheet root.");
            foreach (ReaderCell cell in EnumerateWorksheetCells(worksheet)) {
                ct.ThrowIfCancellationRequested();
                var reference = new XmlCoordinateReference(cell.Reference);
                if (cell.StyleAttribute is string style) {
                    if (OfficeIMO.Core.Internal.OpenXml.OfficeXmlViewValues.ParseUInt32(style) is uint styleIndex) {
                        ValidateCellStyleReference(styleIndex, reference);
                    } else if (style.Length == 0) {
                        ReportEmptyCellStyleRepair(reference);
                    } else {
                        throw CreateInvalidCellStyleException(reference);
                    }
                }
                if (cell.Type == ExcelCellValueType.SharedString) {
                    ValidateSharedStringReference(cell.Value, reference);
                }

                OfficeXmlNode formula = cell.Formula;
                if (formula.Attribute("t") != "shared" || !string.IsNullOrWhiteSpace(formula.InnerText)) {
                    continue;
                }
                if (_opt.UseCachedFormulaResult && cell.Value != null) {
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
            // Qualification reads index/follower text before constructing the XML
            // row cache. Bound those strings too, so preflight cannot bypass the
            // streaming reader's protection with a padded index or formula.
            var textBudget = new XmlDataReaderTextBudget(_opt.MaxXmlDataReaderBufferedCharacters,
                ct.ThrowIfCancellationRequested);
            // Table-backed dimensions can intentionally include empty cells.
            // Otherwise reuse this complete validation scan to discover actual bounds.
            var bounds = _usedRangeA1 == null && (!_hasLiveWorksheetPart || _worksheetPart.FirstPartOfKind(ExcelPartKind.Table).IsNone)
                ? new WorksheetRangeAccumulator() : null;
            var coordinates = bounds != null && Volatile.Read(ref _implicitXmlRowIndexes) == null
                ? new ImplicitXmlRowIndexBuilder() : null;
            bool haveSheetData = false;
            bool completedSheetData = false;
            int sheetDataDepth = -1;
            int rowDepth = -1;
            var worksheetRows = new WorksheetXmlRowSelector();
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                XmlNodeType nodeType = reader.NodeType;
                string localName = reader.LocalName;
                bool isCellElement = worksheetRows.IsCellElement(reader);
                bool isRowElement = localName == "row" && worksheetRows.IsRowElement(reader);
                if (isRowElement) textBudget.Reset();
                if (!isCellElement) {
                    if (bounds != null) {
                        if (nodeType == XmlNodeType.Element && reader.Depth == 0
                            && !SpreadsheetXmlContent.IsSpreadsheetElement(reader, "worksheet")) {
                            bounds = null;
                        } else if (!_hasLiveWorksheetPart && worksheetRows.IsWorksheetChildElement(reader, "tableParts")) {
                            // Table extents may include intentional empty rows and
                            // columns; their discovery retains the SDK owner.
                            bounds = null;
                        } else if (worksheetRows.IsWorksheetChildElement(reader, "sheetData")) {
                            if (haveSheetData) {
                                bounds = null;
                            } else {
                                haveSheetData = true;
                                completedSheetData = reader.IsEmptyElement;
                                sheetDataDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            }
                        } else if (isRowElement) {
                            if (rowDepth >= 0) {
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
                    bounds.AddCell(reference.Text);
                    coordinates?.AddCell(reference.Text);
                }
                XmlStyleAttribute styleIndex = ReadXmlStyleAttribute(reader);
                if (styleIndex.Present) {
                    ValidateCellStyleReference(styleIndex, reference, reader);
                }
                bool sharedStringCell = string.Equals(
                    ReadXmlCellTypeAttribute(reader),
                    "s",
                    StringComparison.Ordinal);
                bool sharedFollower = false;
                bool hasCachedValue = false;
                string? sharedStringReference = null;
                if (reader.IsEmptyElement) {
                    continue;
                }

                int cellDepth = reader.Depth;
                bool advance = true;
                while (advance ? reader.Read() : !reader.EOF) {
                    advance = true;
                    ct.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.EndElement
                        && reader.Depth == cellDepth
                        && reader.LocalName == "c") {
                        break;
                    }
                    if (!IsXmlCellChildElement(reader, cellDepth)) {
                        continue;
                    }

                    if (reader.LocalName == "v") {
                        hasCachedValue = true;
                        if (sharedStringCell) {
                            sharedStringReference = reader.IsEmptyElement
                                ? string.Empty
                                : ReadSimpleElementText(reader, ct, reference, textBudget);
                        } else if (!reader.IsEmptyElement && !ct.CanBeCanceled) {
                            // Only shared-string indexes need their text during the validation pass.
                            reader.Skip();
                            advance = false;
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
                            ReadSimpleElementText(reader, ct, reference, textBudget));
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
            XmlCoordinateReference cellReference,
            XmlDataReaderTextBudget textBudget) {
            int elementDepth = reader.Depth;
            string elementName = reader.LocalName;
            string elementNamespace = reader.NamespaceURI;
            string value;
            try {
                value = textBudget.ReadElementText(reader, advancePastEnd: false);
            } catch (XmlException exception) {
                // Keep a malformed text-only value as a hard validation failure;
                // falling back to the SDK DOM would undo the bounded read.
                throw new InvalidDataException(
                    $"Worksheet cell {cellReference.ToString()} element '{elementName}' must contain only text.", exception);
            }
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

        private void ValidateCellStyleReference(XmlStyleAttribute style, XmlCoordinateReference reference, XmlReader reader) {
            if (style.Valid) {
                ValidateCellStyleReference(style.Index, reference);
                return;
            }

            // The attribute text is read again only here, after parsing has already failed.
            if (reader.GetAttribute("s")?.Length == 0) {
                ReportEmptyCellStyleRepair(reference);
                return;
            }

            throw CreateInvalidCellStyleException(reference);
        }

        private InvalidDataException CreateInvalidCellStyleException(XmlCoordinateReference reference) =>
            new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} contains an invalid cell style index.");

        /// <summary>
        /// Excel recovers an empty style attribute (<c>s=""</c>) as the default style and keeps the cell's value, so
        /// the reader does the same. The first repair on this sheet is reported through
        /// <see cref="ExcelExecutionPolicy.OnInfo"/>.
        /// </summary>
        private void ReportEmptyCellStyleRepair(XmlCoordinateReference reference) {
            if (TryBeginEmptyCellStyleRepairReport()) ReportEmptyCellStyleRepair(reference.ToString());
        }

        /// <summary>Returns true for the first repair on this sheet when a callback will receive it.</summary>
        private bool TryBeginEmptyCellStyleRepairReport() {
            if (_reportedEmptyCellStyleRepair) return false;
            _reportedEmptyCellStyleRepair = true;
            return _opt.Execution.OnInfo != null;
        }

        private void ReportEmptyCellStyleRepair(string cellReference) =>
            _opt.Execution.ReportInfo(
                $"[ReadRepair EmptyCellStyle] Worksheet '{_sheetName}' cell {cellReference} has an empty style attribute; " +
                "its value is read with the default style, as Excel repairs it.");

        private void ValidateCellStyleReference(uint styleIndex, XmlCoordinateReference reference) {
            if (styleIndex < (uint)Styles.CellFormatCount) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} references a missing cell style.");
        }

        private void ValidateSharedStringReference(string? rawIndex, XmlCoordinateReference reference) {
            // A shared-string cell without a value element holds no value, as an untyped empty cell does.
            if (rawIndex == null) return;
            int count = _sst.Count;
            if (TryParseSharedStringIndex(rawIndex, out int index)
                && (uint)index < (uint)count) {
                return;
            }

            throw new InvalidDataException(
                $"Worksheet '{_sheetName}' cell {reference.ToString()} references a missing shared string.");
        }

    }
}
