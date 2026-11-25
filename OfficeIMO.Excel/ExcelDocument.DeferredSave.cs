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
        private readonly List<ExcelDeferredCellEdit> _deferredEdits = new();
        private Dictionary<ExcelDeferredSheet, bool>? _deferredSheetHasTables;

        /// <summary>Records a plain value for the owned save; false when the sheet needs the model for it.</summary>
        internal bool RecordDeferredCellEdit(in ExcelDeferredCellEdit edit) {
            ExcelDeferredSheet? entry = edit.Sheet.DeferredEntry;
            if (entry == null || _deferredWorkbook == null || SheetHasTables(entry)) return false;
            lock (_deferredEditsGate) {
                _deferredEdits.Add(edit);
            }
            MarkPackageDirty();
            MarkFormulaInputMutation();
            return true;
        }

        // A value written to a table's header row renames its column, which the model does.
        private bool SheetHasTables(ExcelDeferredSheet entry) {
            _deferredSheetHasTables ??= new Dictionary<ExcelDeferredSheet, bool>();
            if (_deferredSheetHasTables.TryGetValue(entry, out bool known)) return known;
            bool hasTables = true;
            try {
                using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                hasTables = package.GetRelationships(entry.PartName).Any(static relationship =>
                    string.Equals(relationship.Type, TableRelationshipType, StringComparison.Ordinal));
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException) {
                hasTables = true;
            }
            _deferredSheetHasTables[entry] = hasTables;
            return hasTables;
        }

        /// <summary>Writes the recorded values through the model once the SDK package is open, in their original order.</summary>
        private void ReplayDeferredCellEdits() {
            ExcelDeferredCellEdit[] edits;
            lock (_deferredEditsGate) {
                edits = _deferredEdits.ToArray();
                _deferredEdits.Clear();
            }
            foreach (ExcelDeferredCellEdit edit in edits) edit.Sheet.ReplayDeferredCellEdit(edit);
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
            if (!CanSaveDeferredPackage(options) || !TryCreateDeferredPackageBytes(out byte[] bytes)) return false;
            ThrowIfPackageMaterializationExceedsLimit(bytes.LongLength, options);
            PrepareDestinationStreamForWrite(destination);
            destination.Write(bytes, 0, bytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            CompleteDeferredSave(bytes);
            return true;
        }

        /// <summary>Saves through the owned path into a file whose extension matches the package type.</summary>
        private bool TryWriteDeferredPackageToFile(string path, ExcelSaveOptions? options) {
            if (!CanSaveDeferredPackage(options) || !DeferredPackageTypeMatches(path) || !TryCreateDeferredPackageBytes(out byte[] bytes)) return false;
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
                _deferredEdits.Clear();
            }
            // The caller's stream or file has its own copy; the document keeps the array.
            _deferredPackageBytes = bytes;
            _deferredSheetHasTables = null;
            MarkPackageClean(bytes);
            ReleaseReadSnapshot();
            LastSaveDiagnostics = ExcelSaveDiagnostics.OwnedPackage();
        }

        /// <summary>Builds the saved package, or returns false when something needs the model.</summary>
        private bool TryCreateDeferredPackageBytes(out byte[] bytes) {
            bytes = Array.Empty<byte>();
            ExcelDeferredWorkbook workbook = _deferredWorkbook!;
            ExcelDeferredCellEdit[] edits;
            lock (_deferredEditsGate) {
                edits = _deferredEdits.ToArray();
            }

            try {
                string? sharedStringsPart, stylesPart, corePart;
                using (var source = new MemoryStream(_deferredPackageBytes!, writable: false))
                using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                    sharedStringsPart = package.GetRelatedPartName(workbook.WorkbookPartName, SharedStringsRelationshipType);
                    stylesPart = package.GetRelatedPartName(workbook.WorkbookPartName, StylesRelationshipType);
                    corePart = package.GetRelatedPartName(null, OfficeOpcPackage.CorePropertiesRelationshipType);
                    if (sharedStringsPart != null && !package.PartExists(sharedStringsPart)) return false;
                    if (stylesPart != null && !package.PartExists(stylesPart)) return false;
                    // The model's save sets the modification time in the core properties, which it creates when missing.
                    if (corePart == null || !package.PartExists(corePart)) return false;
                }

                using var input = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcEditablePackage editable = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
                ExcelDeferredSharedStrings? sharedStrings = ExcelDeferredSharedStrings.Load(editable, sharedStringsPart);
                if (sharedStrings == null) return false;
                ExcelDeferredDateStyles? dateStyles = null;
                bool dateStylesLoaded = false;

                // Values resolve in the order they were written, as the model adds shared strings; the last value wins.
                var sheets = new Dictionary<ExcelDeferredSheet, SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>>>();
                foreach (ExcelDeferredCellEdit edit in edits) {
                    ExcelDeferredSheet entry = edit.Sheet.DeferredEntry ?? throw new InvalidOperationException("A recorded worksheet is bound to the model.");
                    ExcelDeferredCellOutput output;
                    switch (edit.Kind) {
                        case ExcelDeferredCellKind.Text:
                            output = sharedStrings.TryGetOrAdd(edit.Text, ExcelSheet.CellValuePlainStringPromotionSharedStringCount, out int index)
                                ? new ExcelDeferredCellOutput("s", index.ToString(CultureInfo.InvariantCulture), date: false)
                                : new ExcelDeferredCellOutput("str", Utilities.ExcelSanitizer.EncodeXString(edit.Text), date: false);
                            break;
                        case ExcelDeferredCellKind.Number:
                            output = new ExcelDeferredCellOutput("n", edit.Text, date: false);
                            break;
                        case ExcelDeferredCellKind.Boolean:
                            output = new ExcelDeferredCellOutput("b", edit.Text, date: false);
                            break;
                        default:
                            if (!dateStylesLoaded) {
                                dateStyles = ExcelDeferredDateStyles.Load(editable, stylesPart);
                                dateStylesLoaded = true;
                            }
                            if (dateStyles == null) return false;
                            output = new ExcelDeferredCellOutput("n", edit.Text, date: true);
                            break;
                    }
                    if (!sheets.TryGetValue(entry, out var rows)) sheets[entry] = rows = new SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>>();
                    if (!rows.TryGetValue(edit.Row, out var cells)) rows[edit.Row] = cells = new SortedDictionary<int, ExcelDeferredCellOutput>();
                    cells[edit.Column] = output;
                }

                foreach (var sheet in sheets) {
                    foreach (var row in sheet.Value) {
                        foreach (var cell in row.Value) {
                            if (cell.Value.Type == "s") sharedStrings.AddedReferences++;
                        }
                    }
                    byte[] merged;
                    int mergedLength;
                    int removedSharedReferences;
                    OfficePooledBytes worksheet = editable.ReadPart(sheet.Key.PartName);
                    using (worksheet) {
                        byte[]? result = ExcelDeferredWorksheetWriter.TryMerge(worksheet.Array, worksheet.Length, sheet.Value,
                            baseStyle => dateStyles?.DateStyle(baseStyle), out mergedLength, out removedSharedReferences);
                        if (result == null) return false;
                        merged = result;
                    }
                    sharedStrings.RemovedReferences += removedSharedReferences;
                    editable.ReplacePart(sheet.Key.PartName, merged, mergedLength);
                }

                if (!sharedStrings.TryWrite(editable, workbook.WorkbookPartName)) return false;
                dateStyles?.Write();
                if (!TrySetModifiedTime(editable, corePart)) return false;

                using var output2 = new MemoryStream(_deferredPackageBytes!.Length + 4096);
                editable.Save(output2);
                bytes = output2.ToArray();
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or FormatException or OverflowException) {
                return false;
            }
        }

        // System.IO.Packaging writes package dates in UTC with seven fractional digits.
        private static bool TrySetModifiedTime(OfficeOpcEditablePackage package, string corePart) {
            OfficeXmlTree core = package.GetXml(corePart);
            int modified = core.FindChild(core.Root, OfficeXmlNamespaces.DublinCoreTerms, "modified");
            if (modified == OfficeXmlTree.NoNode) return false;
            core.SetInnerText(modified, DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture));
            return true;
        }
    }
}
