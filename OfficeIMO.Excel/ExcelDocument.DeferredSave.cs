using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Saving a workbook whose SDK package was never opened writes the loaded package again through OfficeIMO's own
    /// Open XML layer: parts nothing changed are copied as stored, edited worksheets are merged with their new values,
    /// the shared strings and styles gain what the values need, and docProps/core.xml gets the modification time a
    /// save through the model gives it. Anything this path does not cover opens the SDK and saves as before.
    /// </summary>
    /// <remarks>
    /// The model's save also runs a workbook preflight that repairs damaged references and adds default views to
    /// every worksheet. This path leaves parts it did not edit as they were read, as Excel does.
    /// </remarks>
    public partial class ExcelDocument {
        private const string SharedStringsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sharedStrings";
        private const string StylesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles";
        private const string TableRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "table";

        private readonly object _deferredEditsGate = new();
        // Sheets with recorded changes, in the order their first change was made.
        private readonly List<ExcelSheet> _deferredEditedSheets = new();
        private const string CellMetadataRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sheetMetadata";
        private bool? _deferredHasCellMetadata;

        /// <summary>Records a plain value for the owned save; false when the sheet needs the model for it.</summary>
        internal bool RecordDeferredCellEdit(ExcelSheet sheet, in ExcelDeferredCellEdit edit) {
            ExcelDeferredSheet? entry = sheet.DeferredEntry;
            // A value written to a table's header row renames its column, which the model does; styles are plain.
            if (entry == null || _deferredWorkbook == null) return false;
            if (edit.Kind != ExcelDeferredCellKind.Style) {
                if (IsTableHeaderCell(entry, edit.Row, edit.Column)) {
                    OpenDeferredSdk("Cell value in a table header row on worksheet '" + entry.Name + "'");
                    return false;
                }
                if (WorkbookHasCellMetadata()) {
                    OpenDeferredSdk("Cell value in a workbook with cell metadata (dynamic arrays or rich values)");
                    return false;
                }
            }
            lock (_deferredEditsGate) {
                ExcelDeferredCellStore? store = entry.Cells;
                if (store == null) {
                    entry.Cells = store = new ExcelDeferredCellStore();
                    _deferredEditedSheets.Add(sheet);
                }
                store.Add(edit);
            }
            if (!_packageDirty) MarkPackageDirty();
            MarkFormulaInputMutation();
            return true;
        }

        // A value written to a table's header row renames its column, which the model does; the header rows of the
        // sheet's tables are read once.
        private bool IsTableHeaderCell(ExcelDeferredSheet entry, int row, int column) {
            if (!entry.TableHeadersRead) {
                entry.TableHeaders = ReadTableHeaders(entry);
                entry.TableHeadersRead = true;
            }
            List<(int Row, int First, int Last)>? headers = entry.TableHeaders;
            // A table the owned reader cannot read counts as covering every cell.
            if (headers == null) return true;
            for (int index = 0; index < headers.Count; index++) {
                (int headerRow, int first, int last) = headers[index];
                if (row == headerRow && column >= first && column <= last) return true;
            }
            return false;
        }

        private List<(int Row, int First, int Last)>? ReadTableHeaders(ExcelDeferredSheet entry) {
            var headers = new List<(int Row, int First, int Last)>();
            try {
                using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(entry.PartName)) {
                    if (!string.Equals(relationship.Type, TableRelationshipType, StringComparison.Ordinal)) continue;
                    if (relationship.IsExternal || relationship.TargetPartName is not { } target || !package.PartExists(target)) return null;
                    OfficeXmlPullReader reader = package.OpenXml(target, out OfficePooledBytes content, processMarkupCompatibility: false);
                    using (content) {
                        while (reader.Read()) {
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                            if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "table")) return null;
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "headerRowCount", out string count) && count == "0") break;
                            if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string reference)
                                || !ExcelA1.TryParseRange(reference, out int top, out int left, out _, out int right)) {
                                return null;
                            }
                            headers.Add((top, left, right));
                            break;
                        }
                    }
                }
                return headers;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException) {
                return null;
            }
        }

        // Dynamic arrays and rich values live in the cell metadata part, which the model reads before a value is written.
        private bool WorkbookHasCellMetadata() {
            if (_deferredHasCellMetadata is { } known) return known;
            bool found = true;
            try {
                using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                found = package.GetRelatedPartName(_deferredWorkbook!.WorkbookPartName, CellMetadataRelationshipType) != null;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException) {
                found = true;
            }
            _deferredHasCellMetadata = found;
            return found;
        }

        /// <summary>Writes the recorded values through the model once the SDK package is open, in their original order.</summary>
        private void ReplayDeferredCellEdits() {
            ExcelSheet[] sheets;
            lock (_deferredEditsGate) {
                sheets = _deferredEditedSheets.ToArray();
                _deferredEditedSheets.Clear();
            }
            // Every handle to a sheet is bound by now; the one that recorded first applies the sheet's changes.
            foreach (ExcelSheet sheet in sheets) {
                ExcelDeferredCellStore? store = sheet.DeferredCellsForReplay;
                sheet.DeferredCellsForReplay = null;
                if (store != null) sheet.ReplayDeferredCells(store);
            }
        }

        private bool CanSaveDeferredPackage(ExcelSaveOptions? options) =>
            _deferredWorkbook != null
            && !_packagePropertiesDirty
            && !_customDocumentPropertiesDirty
            && !HasCalculationSaveWork(options)
            && (options == null || !options.ValidateOpenXml && !options.SafeRepairDefinedNames && !options.SafePreflight
                && options.Format == ExcelFileFormat.Xlsx);

        /// <summary>Saves through the owned path into a stream, or returns false having written nothing.</summary>
        private bool TryWriteDeferredPackageToStream(Stream destination, ExcelSaveOptions? options) {
            if (!CanSaveDeferredPackage(options)) return false;
            if (!TryCreateDeferredPackageBytes(out byte[] bytes, out string? reason)) {
                OpenDeferredSdk("Save: " + reason);
                return false;
            }
            ThrowIfPackageMaterializationExceedsLimit(bytes.LongLength, options);
            PrepareDestinationStreamForWrite(destination);
            destination.Write(bytes, 0, bytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            CompleteDeferredSave(bytes);
            return true;
        }

        /// <summary>Saves through the owned path into a file whose extension matches the package type.</summary>
        private bool TryWriteDeferredPackageToFile(string path, ExcelSaveOptions? options) {
            if (!CanSaveDeferredPackage(options) || !DeferredPackageTypeMatches(path)) return false;
            if (!TryCreateDeferredPackageBytes(out byte[] bytes, out string? reason)) {
                OpenDeferredSdk("Save: " + reason);
                return false;
            }
            ThrowIfPackageMaterializationExceedsLimit(bytes.LongLength, options);
            CommitPreparedPackageToFile(path, bytes);
            CompleteDeferredSave(bytes);
            FilePath = path;
            return true;
        }

        // A save to another extension changes the package type, which the model does.
        private bool DeferredPackageTypeMatches(string path) {
            string extension = Path.GetExtension(path);
            string? type = ReadDeferredMainContentType();
            return type switch {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" => extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase),
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml" => extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase),
                "application/vnd.ms-excel.template.macroEnabled.main+xml" => extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase),
                _ => false
            };
        }

        private string? ReadDeferredMainContentType() {
            using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            return package.GetContentType(_deferredWorkbook!.WorkbookPartName)?.ToLowerInvariant() switch {
                null => null,
                string lowered => WorkbookMainContentTypes.FirstOrDefault(type => type.Equals(lowered, StringComparison.OrdinalIgnoreCase))
            };
        }

        /// <summary>The document now holds what it saved: later reads and saves start from these bytes.</summary>
        private void CompleteDeferredSave(byte[] bytes) {
            lock (_deferredEditsGate) {
                foreach (ExcelSheet sheet in _deferredEditedSheets) {
                    if (sheet.DeferredEntry is { } entry) entry.Cells = null;
                }
                _deferredEditedSheets.Clear();
            }
            // The caller's stream or file has its own copy; the document keeps the array.
            _deferredPackageBytes = bytes;
            foreach (ExcelDeferredSheet entry in _deferredWorkbook?.Sheets ?? new List<ExcelDeferredSheet>()) entry.TableHeadersRead = false;
            _deferredHasCellMetadata = null;
            MarkPackageClean(bytes);
            ReleaseReadSnapshot();
            LastSaveDiagnostics = ExcelSaveDiagnostics.OwnedPackage();
        }

        /// <summary>Builds the saved package, or returns false when something needs the model.</summary>
        private bool TryCreateDeferredPackageBytes(out byte[] bytes, out string? reason) {
            _deferredSaveDeclineReason = null;
            bool created = TryCreateDeferredPackageBytesCore(out bytes);
            reason = _deferredSaveDeclineReason;
            return created;
        }

        private string? _deferredSaveDeclineReason;

        private bool Decline(string reason) {
            _deferredSaveDeclineReason = reason;
            return false;
        }

        private bool TryCreateDeferredPackageBytesCore(out byte[] bytes) {
            bytes = Array.Empty<byte>();
            ExcelDeferredWorkbook workbook = _deferredWorkbook!;
            ExcelSheet[] editedSheets;
            lock (_deferredEditsGate) {
                editedSheets = _deferredEditedSheets.ToArray();
            }

            try {
                string? sharedStringsPart, stylesPart, corePart;
                using (var source = new MemoryStream(_deferredPackageBytes!, writable: false))
                using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                    sharedStringsPart = package.GetRelatedPartName(workbook.WorkbookPartName, SharedStringsRelationshipType);
                    stylesPart = package.GetRelatedPartName(workbook.WorkbookPartName, StylesRelationshipType);
                    corePart = package.GetRelatedPartName(null, OfficeOpcPackage.CorePropertiesRelationshipType);
                    if (sharedStringsPart != null && !package.PartExists(sharedStringsPart)) return Decline("the shared strings relationship has no part");
                    if (stylesPart != null && !package.PartExists(stylesPart)) return Decline("the styles relationship has no part");
                    // The model's save sets the modification time in the core properties, which it creates when missing.
                    if (corePart != null && !package.PartExists(corePart)) return Decline("the core properties relationship has no part");
                    if (corePart == null && package.PartExists(NewCorePropertiesPartName)) return Decline("the core properties part is not related to the package");
                }

                using var input = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcEditablePackage editable = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
                ExcelDeferredSharedStrings? sharedStrings = ExcelDeferredSharedStrings.Load(editable, sharedStringsPart);
                if (sharedStrings == null) return Decline("the shared strings markup needs the model");
                ExcelOwnedStyles? styles = null;
                bool stylesLoaded = false;

                // Shared strings are added in the order values were written, as the model adds them; a cell's last value
                // wins and its style changes apply in turn when the worksheet is merged.
                foreach (ExcelSheet sheet in editedSheets) {
                    ExcelDeferredCellStore store = sheet.DeferredEntry!.Cells!;
                    for (int index = 0; index < store.Count; index++) {
                        ref ExcelDeferredCellEdit edit = ref store[index];
                        if (edit.Kind == ExcelDeferredCellKind.Text) {
                            sharedStrings.TryGetOrAdd(edit.Text, ExcelSheet.CellValuePlainStringPromotionSharedStringCount, out _);
                        } else if (edit.Kind is ExcelDeferredCellKind.Date or ExcelDeferredCellKind.Style && !stylesLoaded) {
                            styles = ExcelOwnedStyles.Load(editable, stylesPart);
                            stylesLoaded = true;
                            if (styles == null) return Decline("the stylesheet needs the model");
                        }
                    }
                }

                ExcelDateSystem dateSystem = DateSystem;
                long changedBytes = 0;
                foreach (ExcelSheet sheet in editedSheets) {
                    ExcelDeferredSheet entry = sheet.DeferredEntry ?? throw new InvalidOperationException("A recorded worksheet is bound to the model.");
                    var cells = new ExcelDeferredSheetCells(entry.Cells!, sharedStrings.IndexOf, date => ExcelDateSystemConverter.ToSerial(date, dateSystem));
                    byte[] merged;
                    int mergedLength;
                    OfficePooledBytes worksheet = editable.ReadPart(entry.PartName);
                    using (worksheet) {
                        byte[]? result = ExcelDeferredWorksheetWriter.TryMerge(worksheet.Array, worksheet.Length, cells,
                            (baseStyle, operation) => styles?.Apply(baseStyle, operation), out mergedLength,
                            out int addedSharedReferences, out int removedSharedReferences, out string? mergeReason);
                        if (result == null) return Decline(entry.Name + ": " + mergeReason);
                        merged = result;
                        changedBytes += mergedLength;
                        sharedStrings.AddedReferences += addedSharedReferences;
                        sharedStrings.RemovedReferences += removedSharedReferences;
                    }
                    editable.ReplacePart(entry.PartName, merged, mergedLength, pooled: true);
                }

                if (!sharedStrings.TryWrite(editable, workbook.WorkbookPartName)) return Decline("the shared strings part cannot take new items");
                styles?.Write();
                if (corePart == null) {
                    CreateCoreProperties(editable);
                } else if (!TrySetModifiedTime(editable, corePart)) {
                    return Decline("the core properties have no modification time to update");
                }

                // Deflated XML is a fraction of its size; the package rarely outgrows its source by more than its new parts.
                using var output2 = new MemoryStream((int)Math.Min(int.MaxValue / 2, _deferredPackageBytes!.Length + changedBytes / 3 + 16 * 1024));
                editable.Save(output2);
                bytes = output2.ToArray();
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or FormatException or OverflowException) {
                return Decline(exception.Message);
            }
        }

        private const string NewCorePropertiesPartName = "/docProps/core.xml";

        private static string ModifiedTimeText() => DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);

        /// <summary>Adds core properties holding the modification time, as a save through the model does for a package without them.</summary>
        private static void CreateCoreProperties(OfficeOpcEditablePackage package) {
            string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
                + " xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
                + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + ModifiedTimeText() + "</dcterms:modified></cp:coreProperties>";
            package.CreateBinaryPart(NewCorePropertiesPartName, "application/vnd.openxmlformats-package.core-properties+xml", System.Text.Encoding.UTF8.GetBytes(xml));
            package.AddRelationship(null, OfficeOpcPackage.CorePropertiesRelationshipType, "docProps/core.xml");
        }

        // System.IO.Packaging writes package dates in UTC with seven fractional digits.
        private static bool TrySetModifiedTime(OfficeOpcEditablePackage package, string? corePart) {
            if (corePart == null) return false;
            OfficeXmlTree core = package.GetXml(corePart);
            int modified = core.FindChild(core.Root, OfficeXmlNamespaces.DublinCoreTerms, "modified");
            if (modified == OfficeXmlTree.NoNode) return false;
            core.SetInnerText(modified, ModifiedTimeText());
            return true;
        }
    }
}
